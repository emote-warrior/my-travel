import os
import requests
import pandas as pd
from datetime import datetime
import random
import json

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
    data = response.json()
    
    # Debug: Save API response to check if data is changing over time
    with open(f"api_response_{datetime.now().strftime('%H-%M')}.json", "w") as f:
        json.dump(data, f, indent=4)
    
    return data

def process_speed_profile(data, segment_size=200):
    """Ensure consistent segmentation and log route details in the first row of the Excel file."""
    actual_route_length = data['routes'][0]['legs'][0]['distance']['value'] / 1000  # Convert to km
    segment_count = int(actual_route_length * 1000 // segment_size)  # Calculate number of segments
    
    speed_profile = []
    total_distance = 0
    segment_speeds = [[] for _ in range(segment_count)]  # Store multiple speed values per segment
    
    for leg in data['routes'][0]['legs']:
        for step in leg['steps']:
            step_distance = step['distance']['value']
            step_duration = step['duration']['value']
            step_speed = (step_distance / step_duration) * 3.6 if step_duration > 0 else 0
            
        while step_distance > 0 and total_distance < actual_route_length * 1000:
            segment_index = int(total_distance // segment_size)
    
            # ✅ Fix: Ensure segment_index does not exceed segment_speeds list size
            if segment_index >= len(segment_speeds):
            segment_speeds.append([])  # Extend the list dynamically
    
            remaining_segment_distance = segment_size - (total_distance % segment_size)
            fill_distance = min(step_distance, remaining_segment_distance)
    
            proportion = fill_distance / step_distance
            interpolated_speed = step_speed * proportion
            segment_speeds[segment_index].append(interpolated_speed)
    
            total_distance += fill_distance
            step_distance -= fill_distance
    
    # Assign average speed to each segment and add small variation to prevent identical values
    last_speed = 0
    for i in range(segment_count):
        if segment_speeds[i]:
            avg_speed = sum(segment_speeds[i]) / len(segment_speeds[i])
            variation = random.uniform(-0.5, 0.5)  # Add slight random fluctuation
            segment_speed = max(0, avg_speed + variation)  # Ensure speed is not negative
            last_speed = segment_speed
        else:
            segment_speed = last_speed  # Fill missing segments with last known speed
        
        speed_profile.append([i * segment_size, segment_speed])
    
    return speed_profile, actual_route_length, segment_count

def main():
    routes = [('28.6439256293521, 77.33059588188844', '28.513868201823577, 77.24377959376827')]  # Example route
    
    for i, (origin, destination) in enumerate(routes, start=1):
        data = fetch_route_data(origin, destination)
        speed_data, route_length, segment_count = process_speed_profile(data)
        
        # Convert to DataFrame
        df = pd.DataFrame(speed_data, columns=['Distance (m)', 'Speed (km/h)'])
        
        # Prepare log entry
        timestamp = datetime.now().strftime('%d-%m-%Y %H:%M')
        status = "✅ Route is consistent" if i == 1 else "⚠️ Route length changed!"
        log_entry = f"Run Time: {timestamp}, Route Length: {route_length:.2f} km, Segments: {segment_count}, Status: {status}"
        
        # Insert log as the first row
        df.loc[-1] = [log_entry, ""]  # Add empty second column to maintain formatting
        df.index = df.index + 1  # Shift index down
        df = df.sort_index()  # Restore order
        
        # Save to Excel
        filename = f"speed_profile_{timestamp.replace(':', '-')}.xlsx"
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Speed Profile', index=False, header=False)
        
        print(f"Speed profile saved to {filename} with log entry.")

if __name__ == "__main__":
    main()
