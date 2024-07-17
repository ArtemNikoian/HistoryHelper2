# HistoryHelper2
This code will help you create a personal history revision book with all the necessary historical information. You can add any events by simply typing their names in the console. The result will look like this:

![Result image](image.png)

## Setup guide:
1. Download the project, unzip the file, and open the HistoryHelper folder, which is inside the HistoryHelper2-main folder, in your editor (e.g. VS Code).
2. Open the terminal and run these lines of code to install all libraries:

   Mac:
   ```
   python3 -m venv venv
   source venv/bin/activate
   pip3 install -r requirements.txt
   brew install python-tk
   ```

   Windows:
   ```
   python3 -m venv venv
   myenv\Scripts\activate
   pip3 install -r requirements.txt
   sudo apt-get install python3-tk
   ```

3. Run the AI2.py file by clicking the run button in VS Code or writing `python3 AI2.py`.
4. Once you run it for the first time, open the new popped-up window and click the Toggle Configuration button.
5. Choose the Excel file path you want to use (If you don't have any other files made by HistoryHelper2, just use the empty one in HistoryHelper2-main/HistoryHelper/History.xlsx) and the OpenAI API key, which you can create by clicking the Get OpenAI API Key button.
6. Click the Update Configuration button to save the changes.
7. You are all set! You can click the Toggle Configuration button again to close the settings.

## Generating events info
To generate information about the event, write events you would like to add to your Excel file in the empty field above the Process Events button and separate different events with a semicolon. Next, click the Process Events button. When the information is generated, check if it is correct, and click Apply Changes when ready.

## Warnings:
Remember that the Excel file structure should be as in the History file; otherwise, the code will not work properly. <br />There should be no empty colons between the events. <br />Do not delete the first and last Null events. <br /> Be careful about the date of the event; it should always be in the format xx.yy.zzzz or xx.yy.zzzz - end date.

## How to use it
This project helped me improve my history grades from a C- to a B+ without having to put in a lot of extra work. I no longer needed to read through numerous presentations filled with useless information. This is how I used it to improve my efficiency: first, I quickly went through the syllabus for our classes without delving too deeply into it. I only looked for events or terms that I didn't know. As soon as I came across something new, I inputted it into the project, and it automatically added it to my file. Then, I reviewed all the events and terms in the file, only focusing on the names I didn't remember. If I came across something I didn't remember, I looked at all the information below and repeated this process until I remembered everything. This method is beneficial because you don't need to read all the extra information written in the books or presentations that you won't need. This focuses specifically on terms you don't know.

## Thank you for visiting the page
I hope this code will help you improve your history grades while decreasing the preparation time, as it did for me. If you like my project, please give it a star so that I know it helped someone. Good luck!
