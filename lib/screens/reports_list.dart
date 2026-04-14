import 'dart:io';
import 'package:flutter/material.dart';
import '../models/report.dart';
import '../services/report_services.dart';

class ReportListScreen extends StatelessWidget {
  const ReportListScreen({super.key});

  @override
  Widget build(BuildContext context) {
    // Get role (admin or student)
    final role = ModalRoute.of(context)?.settings.arguments as String? ?? "student";

    List<Report> reports = ReportService.reports;

    return Scaffold(
      appBar: AppBar(
        title: Text(role == "admin" ? "All Reports" : "My Reports"),
      ),
      body: reports.isEmpty
          ? const Center(
        child: Text(
          "No reports available",
          style: TextStyle(fontSize: 16),
        ),
      )
          : ListView.builder(
        itemCount: reports.length,
        itemBuilder: (context, index) {
          Report report = reports[index];

          return Card(
            margin: const EdgeInsets.all(10),
            child: ListTile(
              leading: report.imagePath.isNotEmpty
                  ? Image.file(
                File(report.imagePath),
                width: 50,
                height: 50,
                fit: BoxFit.cover,
              )
                  : const Icon(Icons.report),

              title: Text(report.title,
                  style: const TextStyle(fontWeight: FontWeight.bold)),
              subtitle: Text(                      "Status: ${report.status}",
                style: TextStyle(
                  color: report.status == "Pending"
                      ? Colors.orange
                      : Colors.green,
                ),
              ),

              trailing: const Icon(Icons.arrow_forward_ios),

              onTap: () {
                Navigator.pushNamed(
                  context,
                  '/details',
                  arguments: report,
                );
              },
            ),
          );
        },
      ),
    );
  }
}