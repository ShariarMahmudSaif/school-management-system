# High Tech School Management System

A comprehensive school management system with a modern Qt desktop interface, real-time Excel data synchronization, and advanced payment tracking.

## Features

### 🎨 Modern UI
- **Commercial Dashboard Design**: Dark theme with crisp borders, amber accents, and professional styling
- **Real-time Data Sync**: File watcher polls Excel file every 800ms for external changes
- **Responsive Navigation**: Sidebar navigation with active state indicators
- **Live Search & Filtering**: Instant table updates with multiple filter criteria

### 👨‍🎓 Student Management
- **Full CRUD Operations**: Add, edit, delete students with validation
- **Advanced Filtering**: Search by name, ID, contact, age; filter by class, section, field
- **Custom Fields**: Configurable additional fields per student
- **Automatic ID Generation**: Sequential student IDs with configurable prefix (default: STU-)
- **Age Tracking**: Automatic age calculation and display

### 👨‍🏫 Teacher Management
- **Complete CRUD**: Add, edit, delete teachers with full data tracking
- **Contact Management**: Primary and secondary contact information
- **Role Assignment**: Teacher role/position tracking
- **Custom Fields**: Extensible teacher metadata
- **Automatic ID Generation**: Sequential teacher IDs with configurable prefix (default: TEA-)

### 💰 Payment Tracking
- **Dual Entity Support**: Track both student fees and teacher salaries
- **Monthly Payment Records**: Year/month-based payment history
- **Amount Tracking**: Individual payment amounts with defaults
- **Status Management**: Paid/Pending status with visual indicators (green=paid, red=pending)
- **Automated Rollover**: Calculates unpaid months up to 24 months back
- **Pending Balance Calculator**: Shows total pending amount and list of unpaid months
- **Advanced Filtering**: Filter by entity type, month/year, payment status
- **Summary Statistics**: Real-time totals for paid/pending counts and amounts
- **Color-Coded Status**: Green cells for paid, red for pending
- **Quick Payment Entry**: Dialog for setting status and amount

### 📊 Dashboard Analytics
- **Key Metrics Cards**: Student count, teacher count, paid/pending counts
- **Payment Status Pie Chart**: Visual breakdown of current month payments
- **Students by Class Bar Chart**: Top 10 classes distribution
- **Real-time Updates**: Automatic refresh when data changes

### 📝 Activity Log
- **Comprehensive Event Tracking**: Logs all add/edit/delete/payment actions
- **Action Filtering**: Filter by action type (add_student, edit_teacher, set_payment, etc.)
- **Search Capability**: Search by entity ID or details
- **Timestamp Records**: Full audit trail with timestamps
- **Last 500 Events**: Recent activity view

### ⚙️ Settings Management
- **ID Prefixes**: Customize student and teacher ID prefixes
- **Default Amounts**: Set default student fee and teacher salary
- **Custom Fields**: Define additional fields for students and teachers (comma-separated)
- **Default Payment Period**: Set default month/year for payment tracking
- **Live Updates**: Changes apply immediately to UI and data structure

### 🔄 Data Persistence
- **Excel Storage**: All data stored in `school_data.xlsx`
- **Automatic Header Repair**: Detects and fixes blank row 1, duplicate headers, empty rows
- **Cache Invalidation**: Forces fresh disk reads after every write
- **Sheet Structure**:
  - **students**: student_id, first_name, last_name, age, class, section, primary_contact, secondary_contact, custom fields
  - **teachers**: teacher_id, first_name, last_name, role, primary_contact, secondary_contact, custom fields
  - **student_payments**: student_id, year, month, status, amount, updated_at
  - **teacher_payments**: teacher_id, year, month, status, amount, updated_at
  - **activity_log**: timestamp, action, entity_type, entity_id, details

### 🛡️ Error Handling
- **Excel Lock Detection**: Friendly error messages for Windows file lock issues
- **Validation**: Input validation on all forms
- **Exception Logging**: All errors logged to `error_log.txt`
- **User-Friendly Alerts**: Clear error messages with hints for common issues

## Installation

### Prerequisites
- Python 3.13.5 (or compatible version)
- Windows OS (tested on Windows)

### Setup
1. **Clone or download** this project
2. **Create virtual environment**:
   ```powershell
   python -m venv .venv
   ```
3. **Activate virtual environment**:
   ```powershell
   .venv\Scripts\Activate.ps1
   ```
4. **Install dependencies**:
   ```powershell
   pip install -r requirements.txt
   ```

## Usage

### Running the Application
```powershell
python main.py
```

### First Launch
On first run, the system will:
- Create `school_data.xlsx` with proper sheet structure
- Create `settings.json` with default configuration
- Create `error_log.txt` for error tracking

### Navigation
- **Dashboard**: Overview with metrics and charts
- **Students**: Manage student records with search and filtering
- **Teachers**: Manage teacher records
- **Payments**: Track student fees and teacher salaries
- **Activity Log**: View all system activity
- **Settings**: Configure system preferences

### Workflow Examples

#### Adding a Student
1. Navigate to **Students** page
2. Click **Add Student**
3. Fill in first name, last name, age, class, section, contacts
4. Add custom field values if configured
5. Click **OK** to save
6. Student appears immediately in table with auto-generated ID

#### Setting a Payment
1. Navigate to **Payments** page
2. Select entity type (Students or Teachers)
3. Choose month and year
4. Click on a person in the table
5. Click **Set Payment**
6. Enter status (Paid/Pending) and amount
7. Click **OK** to save
8. Table updates with color-coded status
9. Pending total recalculates automatically

#### Viewing Pending Payments
1. Navigate to **Payments** page
2. Select entity type and current month/year
3. Click **Filter: Pending** to show only unpaid
4. **Pending Total** column shows cumulative unpaid amount
5. **Pending Months** column lists all unpaid months (up to last 6 displayed)

#### Configuring Custom Fields
1. Navigate to **Settings** page
2. Enter custom field names (comma-separated) for students or teachers
3. Click **Save Settings**
4. System rebuilds Excel sheet with new columns
5. Add/Edit dialogs now include custom fields

### Real-Time Sync
- **Automatic Refresh**: File watcher checks `school_data.xlsx` every 800ms
- **External Edits**: Changes made in Excel are immediately reflected in UI
- **Multi-User**: Close Excel before editing in app (Windows file lock prevention)
- **Cache Invalidation**: Every write triggers fresh disk read

### Data Import/Export
- **Excel Format**: Standard `.xlsx` format for easy migration
- **Manual Editing**: Can edit `school_data.xlsx` directly in Excel (close file before running app)
- **Backup**: Simply copy `school_data.xlsx` for backup

## Configuration Files

### settings.json
```json
{
  "student_id_prefix": "STU-",
  "teacher_id_prefix": "TEA-",
  "student_custom_fields": ["Guardian Name", "Address"],
  "teacher_custom_fields": ["Department", "Qualification"],
  "default_year": 2025,
  "default_month": 1,
  "default_student_fee": 100.0,
  "default_teacher_salary": 3000.0
}
```

### Excel Sheet Structure
**students** sheet columns:
- student_id (primary key)
- first_name, last_name, age, class, section
- primary_contact, secondary_contact
- [custom fields from settings]

**teachers** sheet columns:
- teacher_id (primary key)
- first_name, last_name, role
- primary_contact, secondary_contact
- [custom fields from settings]

**student_payments** / **teacher_payments** columns:
- student_id / teacher_id (foreign key)
- year, month (composite key with ID)
- status (Paid/Pending)
- amount (payment amount)
- updated_at (timestamp)

**activity_log** columns:
- timestamp (ISO 8601)
- action (add_student, edit_teacher, set_payment, etc.)
- entity_type (student, teacher, settings)
- entity_id
- details (description)

## Payment System Details

### Monthly Tracking
- Payments tracked by year + month
- Default amounts from settings (student_fee, teacher_salary)
- Individual amounts can override defaults
- Missing records treated as Pending with default amount

### Rollover Calculation
- `get_pending_months()`: checks last 24 months
- Returns list of unpaid months with amounts
- Calculates total pending across all months
- Displayed in Payments table for easy tracking

### Status Indicators
- **Green cells**: Paid status
- **Red cells**: Pending status
- **Amount column**: Current month amount
- **Pending Total**: Sum of all unpaid months
- **Pending Months**: Last 6 unpaid months (e.g., "Dec 2024, Jan 2025, Feb 2025")

## Technical Architecture

### Technologies
- **PySide6 (Qt6)**: Desktop GUI framework
- **OpenPyXL**: Excel file manipulation
- **Python 3.13.5**: Core language
- **QtCharts**: Dashboard visualizations

### Design Patterns
- **MVC Separation**: Storage layer, UI layer, settings layer
- **Observer Pattern**: File watcher for real-time sync
- **Repository Pattern**: ExcelStore abstraction
- **Factory Pattern**: Auto-generated IDs

### Performance
- **Lazy Loading**: Workbook cache with invalidation
- **Sub-second Latency**: Write → invalidate → refresh < 1s
- **Efficient Filtering**: Proxy models for fast table filtering
- **Minimal Reflows**: Targeted refresh per page

## Troubleshooting

### App Won't Start
- Check Python version: `python --version` (should be 3.13+)
- Verify dependencies: `pip list | findstr PySide6`
- Check error_log.txt for details

### Data Not Saving
- **Excel file open?** Close `school_data.xlsx` in Excel/OneDrive
- **Permission error?** Run as administrator or check file permissions
- **Disk full?** Verify storage space

### UI Not Updating
- Check file watcher: should update every 800ms
- Verify mtime changing: manually edit Excel and watch UI
- Check error_log.txt for exceptions

### Payment Amounts Wrong
- Check settings: default_student_fee, default_teacher_salary
- Verify Excel: payment sheets have "amount" column
- Run header repair: delete row 1 if blank, restart app

## Advanced Features

### Excel Header Repair
The system automatically detects and repairs:
- **Blank Row 1**: If row 1 is empty and row 2 has headers, deletes row 1
- **Duplicate Headers**: Removes extra header rows
- **Empty Rows**: Cleans up empty rows after header
- **Missing Columns**: Appends missing custom field columns

### Cache Invalidation
- `store.invalidate_cache()`: Forces fresh disk read
- Called after every write operation
- Called by file watcher on external changes
- Ensures UI always shows current disk state

### Monthly Rollover
- Payment records checked for last 24 months
- Missing months treated as pending
- Default amounts applied to missing records
- Total pending calculated on-the-fly

## Roadmap / Future Enhancements
- [ ] Automated monthly payment record creation
- [ ] Email/SMS reminders for pending payments
- [ ] Export to PDF/CSV
- [ ] Attendance tracking
- [ ] Grade management
- [ ] Reporting dashboard with date ranges
- [ ] User authentication and role-based access
- [ ] Database backend (PostgreSQL/MySQL) for large schools
- [ ] Cloud sync (OneDrive, Google Drive)

## License
This project is provided as-is for educational and internal use.

## Support
For issues or questions, check `error_log.txt` for detailed error messages.

## Credits
Built with Python, PySide6, and OpenPyXL.
