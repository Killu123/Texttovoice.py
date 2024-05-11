import win32com.client as wincom
    
speak = wincom.Dispatch("SAPI.SpVoice")



while True:
    answer = input("Do you want to run the program? (yes/no): ").lower()
    
    if answer == "yes":
      
        print("Program is running...")
        text = input("Enter text\n")
        speak.Speak(text)
        
    elif answer == "no":
        print("Exiting program...")
        break  
        
    else:
        print("Invalid input. Please enter 'yes' or 'no'.")