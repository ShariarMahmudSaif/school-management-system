# High Tech School Management System

A desktop application to manage student and teacher records, track payments, and maintain activity logs for High Tech School. Built with Python, CustomTkinter, and OpenPyXL.

## Features
- Manage student and teacher profiles with primary and secondary contact information.
- Track tuition and salary payments with status (Paid/Pending).
- Filter students by class and search by name, ID, class, section, or contact.
- View a dashboard with payment statistics.
- Log all actions in an activity log.
- Customize ID prefixes and add custom fields in the Settings tab.

## Prerequisites
- Python 3.8 or higher
- Git (for version control)

## Setup Instructions
1. **Clone the Repository**:
   ```bash
   git clone <repository-url>
   cd school-management-system
   ```

2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the Application**:
   ```bash
   python school_management.py
   ```

## Usage
- **Students/Teachers Tabs**: Add, edit, or delete profiles. Toggle payment status for specific months and years.
- **Dashboard Tab**: View total students/teachers and their payment status for a selected period.
- **Activity Log Tab**: Review all actions (e.g., adding/editing profiles, toggling payments).
- **Settings Tab**: Update ID prefixes or add custom fields for students/teachers.

## Data Storage
- Data is stored in `school_data.xlsx` (excluded from Git via `.gitignore`).
- Settings are saved in `settings.json` (excluded from Git).
- Logs are written to `error_log.txt` (excluded from Git).

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.