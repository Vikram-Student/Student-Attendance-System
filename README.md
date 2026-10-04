# Student Attendance Management System

A web-based Student Attendance Management System designed to manage students, faculty, subjects, and attendance efficiently through role-based access.

## Features

- HOD dashboard and management
- Student management
- Faculty management
- Subject management
- Faculty attendance marking
- Attendance history and editing
- Student attendance dashboard
- Attendance percentage calculation
- Daily, weekly, and monthly reports
- Search and filtering
- CSV, Excel, and PDF exports
- Role-based authentication
- Input validation and duplicate protection
- Responsive user interface

## User Roles

### HOD
- Manage students
- Manage faculty
- Manage subjects
- View attendance
- Generate attendance reports
- Export reports

### Faculty
- View assigned subjects
- Mark attendance
- Edit attendance
- View attendance history

### Student
- View attendance percentage
- View subject-wise attendance
- View attendance history

## Technology Stack

- **Frontend:** HTML, CSS, JavaScript
- **Backend:** Python, Flask
- **Database:** SQLite
- **Reports:** CSV, Excel, PDF
- **Version Control:** Git

## Project Structure

```text
Student-Attendance-System/
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
├── database/
├── routes/
├── static/
│   ├── css/
│   ├── js/
│   └── images/
└── templates/
    ├── login.html
    ├── hod/
    ├── faculty/
    └── student/