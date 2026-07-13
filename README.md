These are instructions for setting up the QR Micro Viva within a Qualtrics survey.

Preparing the Micro Viva
--

1. In Qualtrics, go to Directories, and click on the "Create a list" button. Enter a name for your list (e.g., "My List").

2. Click on the "Upload a file" option, and upload the "QRTemplate.csv" file. These dummy records will be used to create the unique QR code links to the survey.

3. Create a second list, and click on the "Upload a file" option. This time, you should upload a real list of your students. It should contain an Email field and (optionally) FirstName and LastName. Give the list a name (e.g., "My Students").

4. In the main Qualtrics screen, click on the "Create a new project" button. Select "Survey". Give your survey a name (e.g., "My Micro Viva").

5. Under the "How do you want to start your survey", select "Import a QSF" file. Select the "QR_Micro_Viva.qsf" file.

6. In your survey, click on the "Survey Flow" button. Look for the "Branch on Successful Authentication" box. Under "Authenticate Using Contact", select your contact list of students (e.g., "My Students").

7. Next, click on the "Distributions" tab. Choose the "Generate a trackable link for each contact" button.

8. Under the "Generate Links" box, select your contact list (e.g., "My List"). Set the expiry date for the end of the semester. Click on the "Generate Links" button, and download the CSV file containing the links.

9. Go to https://tbednall.pythonanywhere.com. Select the CSV file, and click "Generate & download Word document". This will create the QR codes, to be distributed to students.


Grading the QR Code Assignment
--

1. Once the QR Micro Viva assignment is completed, download the survey response data from Qualtrics.

2. In the "QR Code Scores" Excel spreadsheet, replace the "QR Data" with the survey response data.

3. In the "Students" worksheet, specify the dates in which you are counting the QR codes (these may correspond to the weeks of your academic semester).

4. Drag down any formulas to calculate the scores for all of your students. Implement any manual adjustments to the scoring (e.g., capping the number of contributions per week).

5. Upload the final (total) score into your learning management system.
