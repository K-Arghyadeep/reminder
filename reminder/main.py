import time
import win32com.client as wincom


def text_to_speech(msg: str, speech_rate: float = 0.5, gender: int = 1):
    speak = wincom.Dispatch("SAPI.SpVoice")
    # speaker_gender = ["male", "female"]
    # Enter the index of the list as speaker voice
    if gender < 0 or gender > 1:
        raise Exception("Gender can have either 0 or 1.")
        return
    if not msg:
        raise Exception("Content of msg variable is empty or None.")
        return
    if speech_rate <= 0 or speech_rate >= 5:
        raise Exception("Invalid value for speech_rate[0 < speech_rate < 5].")
        return
    speaker_number = gender
    vcs = speak.GetVoices()
    speak.Voice
    speak.SetVoice(vcs.Item(speaker_number))
    speak.Rate = speech_rate
    for _ in range(50):
        speak.Speak(msg)


if __name__ == "__main__":
    print("What shall I remind you about?")
    text = str(input())
    print("In how many minutes?")
    local_time = float(input())
    local_time = local_time * 60
    time.sleep(local_time)
    text_to_speech(text)

