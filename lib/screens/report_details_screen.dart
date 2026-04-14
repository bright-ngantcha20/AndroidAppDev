import 'dart:io';
import 'package:flutter/material.dart';
import '../models/report.dart';

class ReportDetailsScreen extends StatefulWidget {
  const ReportDetailsScreen({super.key});

  @override
  State<ReportDetailsScreen> createState() =>
      _ReportDetailsScreenState();
}

class _ReportDetailsScreenState extends State<ReportDetailsScreen> {
  @override
  Widget build(BuildContext context) {
    final args = ModalRoute.of(context)?.settings.arguments;
    final Report report = args as Report;

    return Scaffold(
      appBar: AppBar(
        title: const Text("Report Details"),
      ),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(16),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            // Title
            Text(
              report.title,
              style: const TextStyle(
                fontSize: 22,
                fontWeight: FontWeight.bold,
              ),
            ),

            const SizedBox(height: 15),

            // Image
            report.imagePath.isNotEmpty
                ? ClipRRect(
              borderRadius: BorderRadius.circular(10),
              child: Image.file(
                File(report.imagePath),
                height: 200,
                width: double.infinity,
                fit: BoxFit.cover,
              ),
            )
                : const Text("No image available"),

            const SizedBox(height: 15),

            // Description
            const Text(
              "Description:",
              style: TextStyle(fontWeight: FontWeight.bold),
            ),
            const SizedBox(height: 5),
            Text(report.description),

            const SizedBox(height: 20),

            // Status Badge
            const Text(
              "Status:",
              style: TextStyle(fontWeight: FontWeight.bold),
            ),
            const SizedBox(height: 5),

            Container(
              padding:
              const EdgeInsets.symmetric(horizontal: 10, vertical: 5),
              decoration: BoxDecoration(
                color: report.status == "Pending"
                    ? Colors.orange
                    : Colors.green,
                borderRadius: BorderRadius.circular(8),
              ),
              child: Text(
                report.status,
                style: const TextStyle(color: Colors.white),
              ),
            ),

            const SizedBox(height: 30),

            // Button
            SizedBox(
              width: double.infinity,
              child: ElevatedButton.icon(
                onPressed: report.status == "Resolved"
                    ? null
                    : () {
                  setState(() {
                    report.status = "Resolved";
                  });

                  ScaffoldMessenger.of(context).showSnackBar(
                    const SnackBar(
                      content:
                      Text("Report marked as Resolved"),
                    ),
                  );
                },
                icon: const Icon(Icons.check),
                label: const Text("Mark as Resolved"),
              ),
            ),
          ],
        ),
      ),
    );
  }
}