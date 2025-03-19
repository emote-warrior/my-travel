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

def process_speed_profile(data):
    """Extract travel speed profile every 200 meters."""
    speed_profile = []
    total_distance = 0  # Track cumulative distance
    segment_distance = 200  # 200m segments
    segment_time = 0  # Time for each 200m segment
    
    for leg in data['routes'][0]['legs']:
        for step in leg['steps']:
            distance = step['distance']['value']  # in meters
            duration = step['duration']['value']  # in seconds
            
            while distance > 0:
                if total_distance + distance < segment_distance:
                    segment_time += duration
                    total_distance += distance
                    break
                else:
                    ratio = (segment_distance - total_distance) / distance
                    segment_time += ratio * duration
                    avg_speed = (segment_distance / segment_time) * 3.6  # Convert to km/h
                    speed_profile.append([segment_distance, avg_speed])
                    distance -= (segment_distance - total_distance)
                    duration -= (ratio * duration)
                    total_distance = 0
                    segment_time = 0
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
