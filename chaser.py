"""
Design a program that will automate the process of sending laptop chaser mail

"""

import pandas as pd


def main_menu():
    while True:
        print("""
            1. First Chaser
            2. Second Chaser
            3. Final Chaser
            q. Quit
        """) 
            
        # get user's input
        choice = input("Please select your option: ")

        # run corresponding function based on user input
        menu_dict = {
            "1": first_chaser,
            "2": second_chaser,
            "3": third_chaser,
            "q": None}.get(choice,None)            
        
           
        if choice == "q" or choice =="Q":
            break
        elif menu_dict == None:
            print("Try again...")
        else:
            menu_dict()
        

# function to quit the program
def quit_program():
    answer = input('To go back to the Main menu press 1. Press any other key to quit.')
    if answer == '1':
        main_menu()
    else:
        exit()


def get_info():
    # Read the Excel file
    df = pd.read_excel("laptop.xlsx")
    
    # get the bay number
    locker = input('Please Enter Bay Number: ').upper()

    # lapsafe info
    lapsafe = df[df['Locker'] == locker]


    return lapsafe
    
# function to send first chaser email
def first_chaser():
    # locations = ['Parkside Bridge', 'Millenium Point', 'STEAMHouse']
    # get Lapsafe Info
    lapsafe = get_info()
    
    if not lapsafe.empty:

        # index into the first letter of the Bay Number
        if lapsafe.iloc[0]['Locker'][0] == 'A':
            location = 'Parkside Bridge'
        elif lapsafe.iloc[0]['Locker'][0] in ['B', 'L', 'M', 'N', 'O', 'P']:
            location = 'Millenium Point'
        else:
            location = 'STEAMHouse'

        # read message file
        with open("message1.txt", "r") as file:
            content = file.read()
    
        # Assign Variables to placeholders in the text file
        output = content.format(first_name = lapsafe.iloc[0]['First Name'], 
                                 last_name=lapsafe.iloc[0]['Last Name'],
                                 location = location,
                                 locker = lapsafe.iloc[0]['Locker'],
                                 due_date = lapsafe.iloc[0]['Due Date'],
                                 time = lapsafe.iloc[0]['Time']
                              )
        
        # Write the modified content to a new file
        with open("mail_to_send.txt", "w") as file:
            file.write(output)
        
        print(f'E-mail has been sent, check the "mail_to_send.txt" file')

    else:
        print('No Lapsafe Locker Bay!')
    
    # quit program after operation
    quit_program()
        


def second_chaser():
    get_info()
    print('This is the second chase')

    quit_program()

    

def third_chaser():
    get_info()
    print('This is the third chase')

    quit_program()




# start program
main_menu()