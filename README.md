**ColdBot – Automated Cold Calling Voice Agent**

ColdBot is a Python-based voice assistant created for our Final Year Project (FYP) to automate **cold calling** tasks.  
It listens to spoken input, extracts user details (name, gender, email, phone), and stores them in an Excel file for follow-up.

**Features**
 Speech Recognition: Captures details through voice input.  
 Text-to-Speech: Speaks prompts and confirmations.  
 Excel Integration: Automatically creates `user_records.xlsx` and appends new entries with styled headers.  
 Data Cleaning: Normalizes email addresses and phone numbers (e.g., “at sign” → “@”, spoken numbers → digits).  
 Fallback Input: If speech fails, you can type responses.


**Requirements**

Install Python
- Install [Python 3.9+](https://www.python.org/downloads/).  
- During installation, tick **“Add Python to PATH.”**

**Install Dependencies**
Open a terminal (Command Prompt, PowerShell, or Git Bash) and run:
pip install pyttsx3 SpeechRecognition openpyxl pyaudio


**How to Run**

Clone or download this repository:

git clone https://github.com/YOURUSERNAME/FYP_TASKS.git
cd FYP_TASKS


(Or click Code → Download ZIP and extract.)

Plug in and test your microphone.

Run the program:

python voice_agent.py


Speak when prompted. If speech is not recognized, type your answer.

The collected details are saved to user_records.xlsx in the same folder.

**🧪 Testing Example**
Prompt	Example Spoken Answer
What is your full name?	Aisha Sadiqa
What is your gender?	Female
Email address?	aisha sadiqa at the rate gmail dot com
Phone number?	plus nine two three zero one…

ColdBot will normalize and store:

Timestamp | Aisha Sadiqa | Female | aisha.sadiqa@gmail.com | +92301234567

**## 🎤 Testing Your Microphone/Voice**
To quickly test your system’s text-to-speech:
python test_microphone_tts.py


**Future Improvements**

Add a GUI interface for easier use.

Integrate databases instead of Excel for scalability.

Support multiple languages.

Add an analytics dashboard for call performance.

**⚙ Troubleshooting**
Issue	                          Solution
Mic not working                	Check Windows sound settings or try another microphone/headset.

Speech not recognized	          Ensure a stable internet connection.


