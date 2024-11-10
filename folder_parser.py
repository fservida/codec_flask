import os
import uuid
import json
import pandas as pd
import subprocess
import argparse
from datetime import datetime, timedelta
from tqdm import tqdm

def format_timestamp(timestamp_str):
    """Format the timestamp to 'YYYY-MM-DD HH:MM:SS'."""
    try:
        timestamp = datetime.strptime(timestamp_str, '%Y:%m:%d %H:%M:%S')
        return timestamp.strftime('%Y-%m-%d %H:%M:%S')
    except Exception as e:
        print(f"Error formatting timestamp: {e}")
        return None

def format_video_length(seconds):
    """Format video length in 'HH:MM:SS', rounding up to the nearest second."""
    if seconds is None:
        return None
    # Round up to the nearest second
    total_seconds = int(seconds + 0.999)
    video_length = str(timedelta(seconds=total_seconds))
    # Ensure it is in HH:MM:SS format even if under an hour
    if len(video_length.split(":")) == 2:  # If format is MM:SS
        video_length = f"00:{video_length}"
    return video_length

def get_exif_data(filepath):
    """
    Extract metadata using exiftool with -n and -j, preferring DateTimeOriginal, TrackCreateDate,
    MediaCreateDate, CreateDate in that order, and using file modification time as a last resort.
    """
    try:
        # Run exiftool with JSON output, -n for numeric output, and -c for decimal coordinates
        result = subprocess.run(
            ['exiftool', '-j', '-n', '-c', '%.6f', filepath],
            capture_output=True, text=True
        )
        
        # Parse JSON output
        metadata = json.loads(result.stdout)[0] if result.stdout else {}
        
        # Extract relevant metadata
        raw_create_time = (
            metadata.get("DateTimeOriginal") or
            metadata.get("TrackCreateDate") or
            metadata.get("MediaCreateDate") or
            metadata.get("CreateDate")
        )
        
        if raw_create_time:
            exif_create_time = format_timestamp(raw_create_time)
        else:
            # Fallback to file modification time
            exif_create_time = datetime.fromtimestamp(os.path.getmtime(filepath)).strftime('%Y-%m-%d %H:%M:%S')
        
        # Extract GPS data if available
        gps_lat = metadata.get("GPSLatitude")
        gps_long = metadata.get("GPSLongitude")
        
        # Extract video length in seconds, then format it to HH:MM:SS
        video_length_seconds = metadata.get("Duration")
        video_length = format_video_length(video_length_seconds) if video_length_seconds else None
        
        return exif_create_time, gps_lat, gps_long, video_length
    except Exception as e:
        print(f"Could not read EXIF data for {filepath}: {e}")
        return None, None, None, None

def count_files(directory):
    """Count the total number of files in a directory and its subdirectories."""
    file_count = sum(len(files) for _, _, files in os.walk(directory))
    return file_count

def recurse_directory(directory):
    """Recurse through directory, collecting file data with a progress bar."""
    file_data = []
    
    # Count the total files first for progress tracking
    total_files = count_files(directory)
    
    # Traverse files with a progress bar
    with tqdm(total=total_files, desc="Extracting metadata", unit="file") as pbar:
        for root, dirs, files in os.walk(directory):
            for file in files:
                filepath = os.path.join(root, file)
                file_uuid = str(uuid.uuid4())
                
                # Extract metadata
                exif_create_time, exif_gps_lat, exif_gps_long, video_length = get_exif_data(filepath)
                
                # Append to file data
                file_data.append({
                    "UUID": file_uuid,
                    "Filename": file,
                    "Filepath": filepath,
                    "EXIF Create Time": exif_create_time,
                    "EXIF GPS Latitude": exif_gps_lat,
                    "EXIF GPS Longitude": exif_gps_long,
                    "Video Length (HH:MM:SS)": video_length
                })
                
                # Update progress bar
                pbar.update(1)
    
    # Create a DataFrame
    df = pd.DataFrame(file_data)
    return df

def main():
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description="Recurse a directory to collect metadata from image and video files.")
    parser.add_argument('directory', type=str, help="The path to the directory to recurse")
    parser.add_argument('output_filename', type=str, help="The output Excel file name")
    args = parser.parse_args()
    
    # Process the specified directory
    df = recurse_directory(args.directory)
    
    # Export the DataFrame to an Excel file
    df.to_excel(args.output_filename, index=False)
    print(f"Data exported to {args.output_filename}")

if __name__ == '__main__':
    main()
