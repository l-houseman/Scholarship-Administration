# USRA Application and Transcript Protocol 2026
These scripts are used to automate a manual research administration process of naming and sorting large-volume files; creating reviewer assignments; consolidating hiring information; creating award letters and result notifications. This saved ~40 hours of work time in 2026

## Table of Contents
*[Application Sorting, File Naming, and Validation](#application-sorting-file-naming-and-validation)
*[Reviewer Assignment Creation](#reviewer-assignment-creation)
*[Hiring Information Document Naming and Consolidation](hiring-information-document-naming-and-consolidation)
*[Award Document Creation](#award-document-creation)
*[Award Notification Creation](#award-notification-creation)


## Application Sorting, File Naming, and Validation
## Stage 1: Application Sorting and File Naming Protocol
### How to run this protocol (Application Naming and Sorting Script) - Powershell7
**1.	Confirm the three input folders exist and are populated. This step takes three sets of incoming files:**
- USRA 2026 Applications/
  - Student Application Forms
  - Faculty Application Forms
  - Student Transcripts (Non-USask)
And
- Creates one folder per student under an organized root
- Renames each file into a standard naming convention
- Moves each file into the correct student folder
- Produces a CSV log of every action and a “missing documents” report
**2.	Confirm file naming follows the “final underscore student name” rule.**
**3.	Set:**
- $DryRun = $true
**4.	Run script in PowerShell7.**
**5.	Review:** 
- organize log (look for skips)
- missing report (look for gaps)
**6.	If results look correct:** 
- set $DryRun = $false
**7.	Run again to execute moves and folder creation.**
**8.	Archive the CSV log + missing report for traceability.**


#### Details of what this script does:
**File naming prerequisite (critical)**
!The student name must appear after the final underscore _ in the filename.!
Example:
- something_something_Lastname Firstname.pdf
The script’s Get-StudentNameFromBase() function:\
- takes the base filename (no extension)
- finds the last underscore
- treats the text after that underscore as the student name
If a file does not follow this pattern, it is skipped and logged as Skip_NoStudentSuffix.

**Standardized titles used (what files get renamed to)**
The script maps each source folder to a standardized document title prefix:
- Student application forms → 2026 USRA Student Application Form
- Faculty application forms → 2026 USRA Faculty Application Form
- Non USask transcripts → 2026 USRA Non-USask Transcript

**The final filename format becomes:**
{TitlePrefix}_{StudentName}{Extension}
Examples:
- 2026 USRA Student Application Form_Jane Smith.pdf
- 2026 USRA Faculty Application Form_Jane Smith.pdf
- 2026 USRA Non-USask Transcript_Jane Smith.pdf
  
**Collision handling (duplicate filenames)**
If $AllowOverwrite = $false and a target filename already exists, the script does not overwrite. It instead appends:
- (1), (2), etc.
Example:
- 2026 USRA Student Application Form_Jane Smith.pdf
- 2026 USRA Student Application Form_Jane Smith (1).pdf
  
**Execution mode controls**
The script has two major switches:
***Dry run***
- $DryRun = $true
  - no files are moved
  - the script writes a “Plan” row to the log for each file, showing what would happen
- $DryRun = $false
-   the script creates folders as needed and moves files
- 
**Optional subset processing**
- $ProcessOnly = @('Student Transcripts')
  - limits execution to specific source folders by display name
 
**What the script outputs (for audit and troubleshooting)**
Two CSVs are generated in the parent directory of $OrganizedRoot:
***1) Action log***
USRA_2026_organize_log_YYYYMMDD_HHMMSS.csv
Contains, per file:
- timestamp
- source folder + original filename
- extracted student name
- document type/title prefix
- new filename + destination path
- action: Plan, Moved, or a Skip_* reason
***2) Missing document report***
USRA_2026_missing_report_YYYYMMDD_HHMMSS.csv
Contains, per student folder detected:
- whether they have:
  - Student application
  - Faculty application
  - Transcript
- a “Missing” column listing what is absent
- the expected folder path
This report is built from a hashtable ($StudentDocs) that tracks which doc categories were successfully processed for each student.

## Stage 2: Splitting, naming, and sorting USask transcript files (Transcript Splitter) - Python
*Note: Because USask transcript files may contain full legal names that many students do not use, this still requires some manual consolidation (e.g., legal name on transcript is Susan Anna Joyce Double Lastname but student actually uses Susan Lastname in practice; you will need to consolidate “Susan Anna Joyce Double Lastname” file with “Susan Lastname” file). 

**This step takes one consolidated PDF containing many USask transcripts and produces:**
- One output PDF per student
- Named consistently
- Routed into the student’s folder under the organized root (created by Stage 1)
- With a fallback mechanism when a matching student folder is not found
- With preflight and diagnostics CSVs for audit

#### How to run this protocol (USask Transcript SplitNameSort Automation Script.py)
***Required inputs***
You need three inputs (paths set in the CONFIG section of the script):
1.	Consolidated transcript PDF
CONSOLIDATED_PDF = ...\2026 Tri-Agency USRA USask Transcripts.pdf
2.	Organized root (the student folders created in Step A)
ORGANIZED_ROOT = ...\USRA 2026_Organized
3.	Excel “roster” file mapping student number → first/last name
ROSTER_XLSX = ...\2026 Tri-Agency USRA USask Student Transcript Info.xlsx

**Confirm Step A is completed and ORGANIZED_ROOT exists with student folders.**
1.	Confirm CONSOLIDATED_PDF exists at the configured path.
2.	Confirm ROSTER_XLSX exists and contains student number + first/last name columns.
3.	Set in CONFIG:
   - PREFLIGHT_ONLY = True
Run script.
4.	Review:
  - preflight report
  - diagnostics CSVs
  - confirm folder existence issues
5.	Set:
  - PREFLIGHT_ONLY = False
  - DRY_RUN = True
6.	Run again to see planned outputs in console + log.
7.	Set: 
  - DRY_RUN = False
8.	Run to write per student transcript PDFs.
9.	Review USRA_2026_split_log_*.csv and confirm: 
  - counts align with expected students
  - fallback routing volume is reasonable
  - no errors in the log

#### Details of What this script does:
**Dependencies (local Python packages)**
The script imports:
- pypdf
- pdfplumber
- pandas + openpyxl
- rapidfuzz
  
**How transcript boundaries are detected (the splitting logic)**
The script finds the start of each student transcript by scanning each page’s extracted text for BOTH:
- Student Number: #####
- Name: Last, First
These are detected using regex:
- RE_STUDENT_NUM
- RE_NAME
Every page that contains both patterns is treated as a transcript start page.

**Segment creation**
Once start pages are identified, the script creates segments:
- transcript i starts at start_page[i]
- transcript i ends at start_page[i+1] - 1
- last transcript ends at the final page
  
**Legend handling**
The script treats page 1 of the consolidated PDF as a legend:
- it is excluded from the student transcript content
- it is appended to the end of each student’s output PDF
  
**Canonical naming of students (how final names are chosen)**
The script produces a final student folder/name using this priority order:
1.	Roster mapping by Student Number (preferred)
   - If the student number exists in the Excel roster, use: FirstName LastName from the roster as canonical.
2.	If not in roster:
  - Use the detected First Last from the PDF header.
3.	If canonical names exist and a match is unclear:
  - Use fuzzy matching (rapidfuzz) to pick the closest canonical roster name above a cutoff.
    
**Case normalization behavior**
The script only normalizes capitalization when the extracted name is clearly all caps or all lower. Mixed/intentional casing is preserved.

**Name sanitization**
Names are made filesystem safe:
- illegal filename characters replaced
- extra spaces trimmed
- trailing punctuation removed
  
**Preflight + diagnostics (what happens before writing PDFs)**
Before splitting, the script:
- reads the roster and builds:
  - a student number → canonical name mapping
- scans the PDF to detect transcript start pages
- produces preflight CSVs that document:
  - which student numbers are in the PDF
  - which student numbers are in the roster
  - which canonical folders exist under ORGANIZED_ROOT  
Outputs include:
- preflight_report_TIMESTAMP.csv
- diagnostics_pdf_not_in_excel_TIMESTAMP.csv
- diagnostics_excel_not_in_pdf_TIMESTAMP.csv
  
**PREFLIGHT_ONLY mode**
If PREFLIGHT_ONLY = True, the script runs those checks and exits without writing any PDFs.

**Destination routing rules (where outputs are written)**
For each student transcript segment, the script chooses a destination folder:
  **Normal routing**
  - If a folder exists at: ORGANIZED_ROOT/<Canonical Student Name>/ then output goes there.
  **Missing folder routing (fallback)**
  If the canonical folder does not exist:
- If ROUTE_MISSING_TO_FALLBACK = True, output goes to: ORGANIZED_ROOT/_NO_STUDENT_FOLDER_FOUND/<StudentNumber_CanonicalName>/    

This ensures transcripts still get created and nothing is silently dropped
If fallback routing is disabled and REQUIRE_EXISTING_FOLDERS = True, the segment is skipped and logged.

**Output file naming convention**
For each student, the script writes:
{YEAR} USRA USask Transcript_{student}.pdf
Where {student} is either:
Firstname Lastname (normal routing), or
Firstname Lastname_StudentNumber (fallback routing to reduce ambiguity)

File collisions are handled by appending (1), (2), etc.

**Split run log (audit record)**
A CSV log is written under ORGANIZED_ROOT:
- USRA_2026_split_log_TIMESTAMP.csv
Each row captures:
- start/end pages
- student number
- detected name (last, first)
- final canonical name used
- destination folder + output filename
- action: planned / written / fallback / skipped / error
- notes (including “legend appended”)

## Stage 3: Validating that each student folder exists, and has all  required documents
This script verifies that all USRA applicants have a folder and that each folder contains the required application and transcript documents for adjudication.

**It cross checks:**
- An official Excel roster of student names
- Against the organized USRA application folder structure
- And produces a single CSV audit report identifying:
  - missing folders
  - missing documents
  - extra folders not tied to a student in the roster
This script is designed to be repeatable year to year with minimal changes.

**What the script checks:**
For each student listed in the Excel roster, the script verifies:
Required documents (must exist)
- Faculty Application Form
  2026 USRA Faculty Application Form_Firstname Lastname.pdf
- Student Application Form
  2026 USRA Student Application Form_Firstname Lastname.pdf
Transcript rules
- At least one of the following must exist:
  - USask Transcript
    2026 USRA USask Transcript_Firstname Lastname.pdf
  - Non USask Transcript
    2026 USRA Non-USask Transcript_Firstname Lastname.pdf
- Some students may have both transcripts; this is valid.
  
**Folder level checks**
- Every student in the Excel roster should have exactly one folder
- Folders present on disk but not in the roster are flagged
- The fallback folder
_NO_STUDENT_FOLDER_FOUND
is recorded but not validated, since it is an intentional holding location

**Folder structure assumed**
USRA 2026_Organized\
- ├─ Firstname Lastname\
  - │  ├─ 2026 USRA Faculty Application Form_Firstname Lastname.pdf
  - │  ├─ 2026 USRA Student Application Form_Firstname Lastname.pdf
  - │  ├─ 2026 USRA USask Transcript_Firstname Lastname.pdf
  - │  └─ 2026 USRA Non-USask Transcript_Firstname Lastname.pdf   (optional)
  - │
- ├─ _NO_STUDENT_FOLDER_FOUND\
- │  └─ ...

Document matching is done by filename prefix, not full filename, so minor differences after the student name will not break the script.

**Inputs**
1. Excel roster
An .xlsx file containing at least:
- Student first name
- Student last name
The script is flexible about column headers (e.g., “First Name”, “Student First Name”, etc.).
2. Organized root folder
The directory containing one folder per student.

**Outputs**
The script generates one CSV file in the organized root, named:
USRA_2026_folder_and_file_audit_YYYYMMDD_HHMMSS.csv

**Output columns**
| Column | Meaning |
| ------ | ------- |
| NameOrFolder |	Student name (from Excel) or folder name |
| FolderExists |	YES / NO |
| FacultyApplicationForm |	YES / NO |
| StudentApplicationForm |	YES / NO |
| USaskTranscript |	YES / NO |
| NonUSaskTranscript |	YES / NO |
| Status |	Complete or specific issue(s) |

**Status values you may see**
- Complete
- MissingFolder
- MissingFacultyApplication
- MissingStudentApplication
- MissingTranscript
- ExtraFolder_NotInRoster
- FallbackFolder
Multiple issues are listed as semicolon separated values.

#### How to run the script
1.	Open Command Prompt (or PowerShell).
2.	Navigate to the project directory
3.	Run python Scripts\audit_usra_2026_folders_against_roster.py
4.	Open the generated CSV in Excel and filter by the status column

**Recommended workflow**
1.	Run the transcript splitting script first
2.	Manually resolve items in _NO_STUDENT_FOLDER_FOUND
3.	Run this audit script
4.	Filter the CSV to:
    - Status ≠ Complete
    - Status = MissingFolder
    - Status = ExtraFolder_NotInRoster
5.	Resolve issues and re run until clean

## Reviewer Assignment Creation
Python code based on strict reviewer recusal rules

## Hiring Information Document Naming and Consolidation 
Auto name files with python, consolidate with powerquery, excel formula to sort awarded, alternate list, and unawarded

## Award Document Creation
VBA code to amplify Word MailMerge to create, save as PDF, and name award documents

## Award Notification Creation
VBA code to automate creation of personalized email notifications with unique attachments
