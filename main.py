import pandas as pd
from pytube import YouTube
from moviepy.editor import VideoFileClip
import os
from faster_whisper import WhisperModel
import json

# Function to download the video from a given URL
def download_video(video_id, youtube_url, output_directory):
    try:
        yt = YouTube(youtube_url)
        video_stream = yt.streams.filter(file_extension="mp4").first()
        video_output_filename = f"{video_id}.mp4"
        video_output_path = os.path.join(output_directory, video_output_filename)
        video_stream.download(output_directory, filename=video_output_filename)
        return video_output_path, None  # Return None for error if download is successful
    except Exception as e:
        return None, str(e)  # Return the error message if there's an exception

# Function to convert video to audio
def convert_video_to_audio(video_filename, audio_output_filename):
    video_clip = VideoFileClip(video_filename)
    audio_clip = video_clip.audio
    audio_output_path = os.path.join(os.path.dirname(video_filename), audio_output_filename)
    audio_clip.write_audiofile(audio_output_path)
    video_clip.close()
    audio_clip.close()
    return audio_output_path

# Function to transcribe audio
def transcribe_audio(audio_filename, model):
    segments, _ = model.transcribe(audio_filename, word_timestamps=True, beam_size=5)
    segments = list(segments)
    transcribed_text = []
    word_level_info = []
    for segment in segments:
        for word in segment.words:
            transcribed_text.append(word.word)
            word_level_info.append({'word': word.word, 'start': word.start, 'end': word.end})
    return " ".join(transcribed_text), word_level_info

# Load the Excel file
excel_file_path = "C:/Users/ravin/OneDrive/Documents/YOUTUBEDATA.xlsx"
df = pd.read_excel(excel_file_path)

# Set the output directory
output_directory = "C:/Users/ravin/VSCODE/MODELFILES"

# Set the model parameters
model_size = "large"
model = WhisperModel(model_size, device="cpu", compute_type="int8")

# Find the next available video to download
for index, row in df.iterrows():
    if row['Video Status'] != 'downloaded':
        video_id = row['Video ID']
        youtube_url = row['Video URL']
        video_output_path, error = download_video(video_id, youtube_url, output_directory)

        if error:
            print(f"Error downloading video {video_id}: {error}")
            continue  # Move to the next video if there's an error

        # Convert video to audio
        audio_output_filename = f"{video_id}.mp3"
        audio_output_path = convert_video_to_audio(video_output_path, audio_output_filename)

        # Transcribe the audio
        transcribed_text, word_level_info = transcribe_audio(audio_output_path, model)

        # Update the Excel file with the download and transcription status
        df.at[index, 'Video Status'] = 'downloaded'
        df.at[index, 'Transcription Status'] = 'transcribed'
        df.at[index, 'Transcribed Text'] = transcribed_text

        # Write word-level information to JSON
        json_filename = f"{video_id}.json"
        json_filepath = os.path.join(output_directory, json_filename)
        with open(json_filepath, 'w') as f:
            json.dump(word_level_info, f, indent=4)

        # Update the Excel file with the success status for JSON writing
        df.at[index, 'Json'] = 'yes'

        # Save the updated Excel file
        df.to_excel(excel_file_path, index=False)

        print(f"Video {video_id} downloaded to: {video_output_path}")
        print(f"Audio {audio_output_filename} converted to: {audio_output_path}")
        print("Audio transcribed successfully.")
        print(f"Word-level information saved to: {json_filepath}")
        print("Excel file updated.")
        break  # Exit the loop after downloading the first available video
else:
    print("All videos have already been downloaded.")
