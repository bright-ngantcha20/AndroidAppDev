# campus_report

## Project Overview

The **Campus Report App** is a mobile application developed using **Flutter (Dart)** that allows students to report issues within a campus environment using real-time data capture features such as the **camera** and **microphone**.

The system enables efficient communication between students and administrators by allowing issues to be reported, tracked, and resolved digitally.

---

## Objectives

* To provide a simple platform for reporting campus-related issues
* To utilize mobile device sensors such as:

  * 📷 Camera (for capturing images)
  * 🎤 Microphone (for voice input)
* To allow administrators to monitor and resolve reported issues
* To improve response time and communication within the campus

---

## Users of the System

### 1. Students

* Register and login
* Report issues (with image and voice description)
* View submitted reports
* Track report status

### 2. Admin

* Login as administrator
* View all reports
* Update report status (e.g., Pending → Resolved)

---

## System Architecture

The application follows a simple structure:

* **Presentation Layer** → UI screens (Flutter widgets)
* **Logic Layer** → Handles user actions and navigation
* **Data Layer** → Temporary storage using in-memory list (`ReportService`)

---

## Project Structure

```bash
lib/
│
├── models/
│   └── report.dart
│
├── services/
│   └── report_service.dart
│
├── screens/
│   ├── login_screen.dart
│   ├── home_screen.dart
│   ├── report_screen.dart
│   ├── report_list.dart
│   └── report_details_screen.dart
│
└── main.dart
```

---

## Features Implemented

### Authentication

* User registration and login using local storage (`SharedPreferences`)
* Role-based access (Student / Admin)

---

### Dashboard

* Different home screens for:

  * Students
  * Admins

---

### Report Submission

* Capture image using device camera
* Input description via text or voice
* Submit report

---

### Voice Input

* Converts speech to text using `speech_to_text` package

---

### Report Management

* View list of reports
* View detailed report information
* Update report status (Admin only)

---

## Technologies Used

* **Flutter (Dart)**
* **Android Studio**
* **SharedPreferences** (local storage)
* **Image Picker** (camera access)
* **Speech-to-Text** (voice input)
* **Permission Handler** (runtime permissions)

---

## Android Configuration

The app uses the following permissions:

```xml
<uses-permission android:name="android.permission.CAMERA"/>
<uses-permission android:name="android.permission.RECORD_AUDIO"/>
```

---

## ⚠️ Limitations

* Data is stored temporarily (not persistent)
* No real backend (no Firebase/Database yet)
* User authentication is basic

---

## 🔮 Future Improvements

* Integration with Firebase or SQLite for persistent storage
* Push notifications for report updates
* User profile management
* Location tracking for reported issues
* Improved UI/UX design

---

## 📚 Conclusion

The Campus Report App demonstrates how mobile technologies and sensors can be used to improve communication and problem reporting within a campus environment. It provides a practical implementation of real-time data collection and role-based system design.

---
