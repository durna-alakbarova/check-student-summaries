# check-student-summaries

###### README ######

> *I created this project during my work as the human subject pool coordinator. Students in the subject pool would submit summaries of journal articles that I had to grade. Frequently, there were over 100 submissions in a cycle, and I wrote this code to automate checking whether the summary submissions were completed on articles from the latest volume and issue of the appropriate journals.
>
> This project showcases my programming skills in R including getting user input using dialog boxes, my ability to scrape data from webpages, and how I utilize programming to work with Microsoft Office programs, which are more accessible to users.*




*** This project is a showcase of my solution  for a specific project and was not created specifically with reusability in mind ***
 
# What the project does:
This project extracts data from journal websites and processes student paper summaries. The journal websites are dynamic pages that are updated whenever a new issue is published. The code collects journal's current volume and issue numbers and a list of the article titles in the publication. It then compares this information with the information inputted by students, and generates a master Excel sheet that contains information from student submission across several submission cycles. The code then creates individual Word documents for each student's summary for easy reading, numbering them when there is more than one submission.

# Why the project is useful:
 This project automates the process of gathering journal information and validating student submissions, saving time and reducing manual errors. It also organizes the data in a structured format for easy review and grading.

# How users can get started with the project:
 1. Ensure all required R packages are installed.
 2. Make sure Java is installed on your computer: https://www.oracle.com/java/technologies/downloads/
 3. Update the `currentSemester` and `journalURLs` variables with the relevant information.
 4. Run the script to extract journal data, process student submissions, and generate the necessary files.

## Data Statement
The sample data provided in the src folder is generated using ChatGPT and contains no data from real people
