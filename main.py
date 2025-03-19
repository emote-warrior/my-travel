import os
import requests
import pandas as pd
from datetime import datetime
import random

# Google API Key
API_KEY = os.getenv('GOOGLE_API_KEY')

def fetch_route_data(origin, destination):
    """Fetch detailed route data including steps, forcing recalculation by using a timestamp."""
    params = {
        'origin': origin,
        'destination': destination,
        'key': API_KEY,
        'departure_time': int(datetime.now().timestamp()),  # Forces Google to recalculate traffic
        'traffic_model': 'best_guess',
        'alternatives': 'false',
        'steps': 'true'
    }
    response = requests.get('https://maps.googleapis.com/maps/api/directions/json', params=params)
    return response.json()

def process_speed_profile(data, segment_size=200):
    """Ensure same number of segments each time for the entire route with 200m segments."""
    
    actual_route_length = data['routes'][0]['legs'][0]['distance']['value']  # Get actual distance
    segment_count = actual_route_length // segment_size  # Consistent segmentation
    
    speed_profile = []
    total_distance = 0  # Track cumulative distance
    segment_speeds = [None] * segment_count  # Pre-allocate slots for consistency
    
    for leg in data['routes'][0]['legs']:
        for step in leg['steps']:
            step_distance = step['distance']['value']  # in meters
            step_duration = step['duration']['value']  # in seconds
            step_speed = (step_distance / step_duration) * 3.6 if step_duration > 0 else 0  # km/h
            
            while step_distance > 0 and total_distance < actual_route_length:
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
    # Add a tiny random offset to prevent caching issues
    lat_offset = random.uniform(-0.0001, 0.0001)  # Tiny shift (~10m)
    lng_offset = random.uniform(-0.0001, 0.0001)
    origin = f"{28.6439256293521 + lat_offset}, {77.33059588188844 + lng_offset}"
    destination = f"{28.513868201823577 + lat_offset}, {77.24377959376827 + lng_offset}"
    
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
