import os
import requests
import pandas as pd
from datetime import datetime

# Google API Key
API_KEY = os.getenv('GOOGLE_API_KEY')

def fetch_route_data(origin, destination):
    """Fetch detailed route data including steps."""
    params = {
        'origin': origin,
        'destination': destination,
        'key': API_KEY,
        'departure_time': 'now',
        'traffic_model': 'best_guess',
        'alternatives': 'false',
        'steps': 'true'
    }
    response = requests.get('https://maps.googleapis.com/maps/api/directions/json', params=params)
    return response.json()

def process_speed_profile(data, route_length=25000, segment_size=200):
    """Ensure exactly 125 segments of 200m each over the full route."""
    speed_profile = []
    total_distance = 0  # Track cumulative distance
    segment_time = 0  # Time for each 200m segment
    segment_count = route_length // segment_size  # 125 segments for 25km
    segment_speeds = [None] * segment_count  # Pre-allocate 125 slots
    
    for leg in data['routes'][0]['legs']:
        for step in leg['steps']:
            step_distance = step['distance']['value']  # in meters
            step_duration = step['duration']['value']  # in seconds
            step_speed = (step_distance / step_duration) * 3.6 if step_duration > 0 else 0  # km/h
            
            while step_distance > 0 and total_distance < route_length:
                segment_index = total_distance // segment_size
                remaining_segment_distance = segment_size - (total_distance % segment_size)
                
                if step_distance >= remaining_segment_distance:
                    # Fill the current segment with interpolated speed
                    segment_speeds[segment_index] = step_speed
                    total_distance += remaining_segment_distance
                    step_distance -= remaining_segment_distance
                    step_duration -= (remaining_segment_distance / step_speed) * 3600 if step_speed > 0 else 0
                else:
                    total_distance += step_distance
                    step_distance = 0
    
    # Assign default speed for missing segments (use last known speed or 0)
    last_speed = 0
    for i in range(segment_count):
        if segment_speeds[i] is None:
            segment_speeds[i] = last_speed
        else:
            last_speed = segment_speeds[i]
        speed_profile.append([i * segment_size, segment_speeds[i]])
    
    return speed_profile

def main():
    routes = [('28.6439256293521, 77.33059588188844', '28.513868201823577, 77.24377959376827')]  # Example route
    
    for i, (origin, destination) in enumerate(routes, start=1):
        data = fetch_route_data(origin, destination)
        speed_data = process_speed_profile(data)
        
        # Convert to DataFrame
        df = pd.DataFrame(speed_data, columns=['Distance (m)', 'Speed (km/h)'])
        
        # Save to Excel with formatting
        timestamp = datetime.now().strftime('%d-%m-%Y %H-%M')
        filename = f"speed_profile_{timestamp}.xlsx"
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Speed Profile', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Speed Profile']
            for col in worksheet.columns:
                max_length = max(len(str(cell.value)) for cell in col)
                worksheet.column_dimensions[col[0].column_letter].width = max_length + 2
        
        print(f"Speed profile saved to {filename}")

if __name__ == "__main__":
    main()
