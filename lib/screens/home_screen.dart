import 'package:flutter/material.dart';

class HomeScreen extends StatelessWidget {
  const HomeScreen({super.key});

  @override
  Widget build(BuildContext context) {
    // Safer role handling
    final role =
        ModalRoute.of(context)?.settings.arguments as String? ?? "student";

    return Scaffold(
      appBar: AppBar(
        title: const Text("Campus Report"),
        centerTitle: true,

        // 🔴 Logout button
        actions: [
          IconButton(
            icon: const Icon(Icons.logout),
            onPressed: () {
              Navigator.pushReplacementNamed(context, '/');
            },
          ),
        ],
      ),

      body: role == "admin"
          ? const AdminHomeLayout()
          : StudentHomeLayout(role: role), // pass role
    );
  }
}

// ================= STUDENT VIEW =================
class StudentHomeLayout extends StatelessWidget {
  final String role;

  const StudentHomeLayout({super.key, required this.role});

  @override
  Widget build(BuildContext context) {
    return ListView(
      padding: const EdgeInsets.all(16),
      children: [
        // ✅ Using interpolation
        Text(
          "Welcome, $role",
          style: const TextStyle(fontSize: 18),
        ),

        const SizedBox(height: 10),

        const Text(
          "Student Dashboard",
          style: TextStyle(fontSize: 22, fontWeight: FontWeight.bold),
        ),

        const SizedBox(height: 30),

        ElevatedButton.icon(
          onPressed: () {
            Navigator.pushNamed(context, '/report');
          },
          icon: const Icon(Icons.report),
          label: const Text("Report an Issue"),
        ),

        const SizedBox(height: 15),

        ElevatedButton.icon(
          onPressed: () {
            Navigator.pushNamed(
              context,
              '/reports',
              arguments: "student",
            );
          },
          icon: const Icon(Icons.list),
          label: const Text("View My Reports"),
        ),
      ],
    );
  }
}

// ================= ADMIN VIEW =================
class AdminHomeLayout extends StatelessWidget {
  const AdminHomeLayout({super.key});

  @override
  Widget build(BuildContext context) {
    return ListView(
      padding: const EdgeInsets.all(16),
      children: [
        const Text(
          "Admin Dashboard",
          style: TextStyle(fontSize: 22, fontWeight: FontWeight.bold),
        ),

        const SizedBox(height: 30),

        ElevatedButton.icon(
          onPressed: () {
            Navigator.pushNamed(
              context,
              '/reports',
              arguments: "admin",
            );
          },
          icon: const Icon(Icons.admin_panel_settings),
          label: const Text("View All Reports"),
        ),
      ],
    );
  }
}