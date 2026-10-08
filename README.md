#  Office Meeting Room Booking System

A simple desktop-based **Office Meeting Room Booking System** developed using **Python, Tkinter, and Excel**. The application helps employees manage meeting-room bookings through an easy-to-use graphical interface.

##  Project Overview

The Office Meeting Room Booking System allows users to book and manage meeting rooms in an office environment.

The system provides a login screen and dashboard where users can:

- Book a meeting room
- View all bookings
- Update an existing booking
- Cancel a booking
- Search bookings by employee name
- Store booking information in an Excel file

The application uses **Tkinter** for the graphical user interface and **OpenPyXL** for storing and managing booking data in an Excel workbook.

##  Features

###  Login System
- Admin login authentication
- Displays an error message for invalid credentials
- Provides access to the main dashboard after successful login

### Book Meeting Room
Users can enter:

- Booking ID
- Employee Name
- Department
- Room Name
- Meeting Date
- Start Time
- End Time
- Number of People
- Meeting Purpose

The system also checks whether the Booking ID already exists before saving a new booking.

### 👀 View Bookings
Displays all existing meeting-room bookings in a table format using Tkinter's `Treeview`.

###  Update Booking
Users can enter a Booking ID and update booking information such as:

- Employee Name
- Department
- Room Name
- Meeting Date
- Start Time
- End Time
- Number of People
- Purpose

### Cancel Booking
Users can cancel an existing booking by entering its Booking ID.

###  Search Booking
Users can search for bookings using the employee's name.

##  Technologies Used

| Technology | Purpose |
|---|---|
| Python | Application development |
| Tkinter | Graphical User Interface |
| OpenPyXL | Excel file handling |
| Excel | Booking data storage |

## 📂 Project Structure

```text
Office-Meeting-Room-Booking-System/
│
├── python project 3.py
└── meeting_room_bookings.xlsx
```

> `meeting_room_bookings.xlsx` is created automatically by the application if it does not already exist.

## ⚙️ Installation

### 1. Clone the Repository

```bash
git clone https://github.com/Tejaswini-Deore/Office-Meeting-Room-Booking-System.git
```

### 2. Open the Project

```bash
cd Office-Meeting-Room-Booking-System
```

### 3. Install Required Library

Install OpenPyXL using:

```bash
pip install openpyxl
```

Tkinter is generally included with standard Python installations on Windows.

### 4. Run the Application

```bash
python "python project 3.py"
```

## 🔑 Login Credentials

For the current version of the application:

```text
Username: admin
Password: 1234
```



The application creates an Excel workbook named:

```text
meeting_room_bookings.xlsx
```

The workbook stores the following information:

```text
Booking ID
Employee Name
Department
Room Name
Meeting Date
Start Time
End Time
Number of People
Purpose
```

## Application Workflow

```text
Start Application
       ↓
    Login
       ↓
   Dashboard
       ↓
 ┌─────┼──────────────┐
 ↓     ↓              ↓
Book  View          Search
 ↓     ↓              ↓
Save  Display       Find
 ↓
Excel File
       ↓
Update / Cancel Booking
```



