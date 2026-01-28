import speech_recognition as sr
import win32com.client
import requests
import json
import cv2
import pytesseract
import datetime

# ================= SPEAKER =================
speaker = win32com.client.Dispatch("SAPI.SpVoice")

def speak(text):
    print("\nAssistant:", text)
    speaker.Speak(text)

# ================= MICROPHONE =================
recognizer = sr.Recognizer()
recognizer.energy_threshold = 300
recognizer.dynamic_energy_threshold = True

MIC_INDEX = None  # Auto-detect default mic

def listen():
    with sr.Microphone(device_index=MIC_INDEX) as source:
        print("\n🎤 Listening...")
        recognizer.adjust_for_ambient_noise(source, duration=1)
        audio = recognizer.listen(source)

    try:
        text = recognizer.recognize_google(audio)
        print("You:", text)
        return text.lower()
    except sr.UnknownValueError:
        speak("Sorry, I did not understand")
        return ""
    except sr.RequestError:
        speak("Speech service error")
        return ""

# ================= OLLAMA PHI =================
OLLAMA_URL = "http://localhost:11434/api/generate"
MODEL_NAME = "phi"

def ask_phi(prompt):
    payload = {
        "model": MODEL_NAME,
        "prompt": prompt,
        "stream": False
    }

    try:
        response = requests.post(
            OLLAMA_URL,
            headers={"Content-Type": "application/json"},
            data=json.dumps(payload)
        )

        if response.status_code == 200:
            return response.json()["response"]
        else:
            return "AI error occurred"
    except requests.exceptions.ConnectionError:
        return "Ollama is not running"

# ================= DATE & TIME =================
def handle_date_time(command):
    now = datetime.datetime.now()

    if "time" in command:
        return "The current time is " + now.strftime("%I %M %p")

    if "date" in command or "day" in command:
        return "Today is " + now.strftime("%A, %d %B %Y")

    return None

# ================= CAMERA OCR =================
def read_book_with_camera():
    speak("Opening camera. Show the book and press S to read. Press Q to exit.")

    cap = cv2.VideoCapture(0, cv2.CAP_ANY)

    if not cap.isOpened():
        speak("Camera not accessible")
        return

    while True:
        ret, frame = cap.read()
        if not ret:
            speak("Camera error")
            break

        cv2.imshow("Book Reader | S = Scan | Q = Exit", frame)
        key = cv2.waitKey(1) & 0xFF

        if key == ord('s'):
            speak("Reading the book")

            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            gray = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)[1]

            text = pytesseract.image_to_string(gray)

            if text.strip() == "":
                speak("No readable text found")
            else:
                print("\n--- OCR TEXT ---\n")
                print(text)
                speak(text)

        elif key == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

# ================= MAIN LOOP =================
speak("Smart Student Assistant started.")
speak("You can speak now.")

while True:
    user_input = listen()

    if user_input == "":
        continue

    if any(word in user_input for word in ["exit", "stop", "bye"]):
        speak("Goodbye. Have a nice day.")
        break

    if "read" in user_input:
        read_book_with_camera()
        speak("Reading completed")
        continue

    # ---- DATE / TIME ----
    dt_reply = handle_date_time(user_input)
    if dt_reply:
        speak(dt_reply)
        continue

    # ---- AI ----
    ai_reply = ask_phi(user_input)
    speak(ai_reply)
