import pyaudio
import speech_recognition as sr
import openpyxl

# Set up PyAudio for live audio streaming
def live_audio_transcription(output_file="transcriptions.xlsx"):
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    # Set up Excel file
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Transcriptions"
    sheet.append(["Timestamp", "Transcription"])

    print("Listening... (Press Ctrl+C to stop)")
    try:
        while True:
            with mic as source:
                recognizer.adjust_for_ambient_noise(source)
                print("Listening for input...")
                audio = recognizer.listen(source)

            try:
                # Use Google Speech Recognition to transcribe audio
                transcription = recognizer.recognize_google(audio)
                print(f"Transcription: {transcription}")

                # Save transcription to Excel
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                sheet.append([timestamp, transcription])
                workbook.save(output_file)
                print(f"Saved transcription to {output_file}")

            except sr.UnknownValueError:
                print("Google Speech Recognition could not understand audio")
            except sr.RequestError as e:
                print(f"Could not request results from Google Speech Recognition service; {e}")
    except KeyboardInterrupt:
        print("Stopping transcription...")
        workbook.save(output_file)

# Call the function
live_audio_transcription()
