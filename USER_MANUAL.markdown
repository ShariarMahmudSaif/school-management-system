# High Tech School Management System User Manual

This manual guides you through setting up the High Tech School Management System locally and uploading it to GitHub using Git commands.

## Prerequisites
- **Python 3.8 or higher**: Download from [python.org](https://www.python.org/downloads/).
- **Git**: Install from [git-scm.com](https://git-scm.com/downloads).
- **GitHub Account**: Sign up at [github.com](https://github.com).
- **Text Editor**: Any editor like VS Code, PyCharm, or Notepad.

## Step 1: Set Up the Project Locally
1. **Create a Project Directory**:
   - Create a folder named `school-management-system` on your computer (e.g., `C:\Projects\school-management-system` on Windows or `~/Projects/school-management-system` on macOS/Linux).
   - Copy the following files into this folder:
     - `school_management.py`
     - `.gitignore`
     - `requirements.txt`
     - `README.md`
     - `LICENSE`
     - `USER_MANUAL.md` (this file)

2. **Install Dependencies**:
   - Open a terminal (Command Prompt on Windows, Terminal on macOS/Linux).
   - Navigate to the project directory:
     ```bash
     cd path/to/school-management-system
     ```
   - Install required Python libraries:
     ```bash
     pip install -r requirements.txt
     ```

3. **Test the Application**:
   - Run the application to ensure it works:
     ```bash
     python school_management.py
     ```
   - The application should open a window with tabs for Students, Teachers, Dashboard, Activity Log, and Settings.
   - Test adding a student or teacher, toggling payments, and checking the dashboard.

## Step 2: Initialize a Git Repository
1. **Initialize Git**:
   - In the terminal, navigate to the project directory:
     ```bash
     cd path/to/school-management-system
     ```
   - Initialize a new Git repository:
     ```bash
     git init
     ```

2. **Add Files to Git**:
   - Stage all project files:
     ```bash
     git add .
     ```
   - The `.gitignore` file ensures that `school_data.xlsx`, `settings.json`, and `error_log.txt` are not tracked.

3. **Commit the Files**:
   - Create an initial commit:
     ```bash
     git commit -m "Initial commit of school management system"
     ```

## Step 3: Create a GitHub Repository
1. **Log in to GitHub**:
   - Go to [github.com](https://github.com) and sign in.

2. **Create a New Repository**:
   - Click the "+" icon in the top-right corner and select "New repository."
   - Enter a repository name (e.g., `school-management-system`).
   - Choose "Public" (or "Private" if preferred).
   - **Do not** initialize with a README, .gitignore, or license (these are already in your project).
   - Click "Create repository."

3. **Copy the Repository URL**:
   - After creating the repository, copy the HTTPS URL (e.g., `https://github.com/your-username/school-management-system.git`).

## Step 4: Push to GitHub
1. **Link Local Repository to GitHub**:
   - In the terminal, add the GitHub repository as a remote:
     ```bash
     git remote add origin https://github.com/your-username/school-management-system.git
     ```

2. **Push the Code**:
   - Push your local repository to GitHub:
     ```bash
     git push -u origin main
     ```
   - If prompted, enter your GitHub username and password (or use a personal access token).

3. **Verify on GitHub**:
   - Visit your GitHub repository URL in a browser.
   - Confirm that all files (`school_management.py`, `.gitignore`, `requirements.txt`, `README.md`, `LICENSE`, `USER_MANUAL.md`) are uploaded.

## Step 5: Using the Application
- **Students/Teachers Tabs**:
  - Click "Add Student" or "Add Teacher" to create a new profile.
  - Enter details like ID (auto-generated), Name, Class/Section (for students), Session Year (for teachers), Primary Contact, Secondary Contact (optional), and Tuition/Salary Amount.
  - Use "Edit" or "Delete" after selecting a profile by clicking its card.
  - Toggle payment status using the "Toggle Payment" button.
- **Dashboard Tab**:
  - View total students/teachers and their payment status for a selected month/year.
- **Activity Log Tab**:
  - Review all actions (e.g., adding, editing, deleting profiles, toggling payments).
- **Settings Tab**:
  - Update ID prefixes (e.g., STU-, TCH-).
  - Add custom fields for students or teachers.

## Troubleshooting
- **Git Errors**:
  - If `git push` fails, ensure you have internet access and correct credentials. Use `git remote -v` to verify the remote URL.
- **Application Errors**:
  - Check `error_log.txt` in the project directory for detailed error messages.
  - Ensure all dependencies are installed (`pip install -r requirements.txt`).
- **Excel File Issues**:
  - If `school_data.xlsx` is not created, ensure you have write permissions in the project directory.

## Contributing
- Fork the repository on GitHub.
- Make changes locally, commit, and push to your fork.
- Create a pull request to the original repository.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.