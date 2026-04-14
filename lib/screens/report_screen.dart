import 'dart:io';
import 'package:flutter/material.dart';
import 'package:image_picker/image_picker.dart';
import 'package:speech_to_text/speech_to_text.dart';
import 'package:permission_handler/permission_handler.dart';

import '../models/report.dart';
import '../services/report_services.dart';

class ReportScreen extends StatefulWidget {
  const ReportScreen({super.key});

  @override
  State<ReportScreen> createState() => _ReportScreenState();
}

class _ReportScreenState extends State<ReportScreen> {
  final TextEditingController titleController = TextEditingController();
  final TextEditingController descriptionController = TextEditingController();

  File? _image;
  final ImagePicker _picker = ImagePicker();

  final SpeechToText _speech = SpeechToText();
  bool _isListening = false;
  bool _isSubmitting = false;

  // 📷 Capture Image
  Future<void> _pickImage() async {
    var status = await Permission.camera.request();

    if (!status.isGranted) return;

    final pickedFile = await _picker.pickImage(source: ImageSource.camera);

    if (pickedFile != null && mounted) {
      setState(() {
        _image = File(pickedFile.path);
      });
    }
  }

  // 🎤 Start Voice Input
  Future<void> _startListening() async {
    var status = await Permission.microphone.request();

    if (!status.isGranted) return;

    bool available = await _speech.initialize();

    if (available && mounted) {
      setState(() => _isListening = true);

      _speech.listen(onResult: (result) {
        if (!mounted) return;

        setState(() {
          descriptionController.text = result.recognizedWords;
        });
      });
    }
  }

  // 🎤 Stop Listening
  void _stopListening() {
    _speech.stop();
    if (!mounted) return;
    setState(() => _isListening = false);
  }

  // 📤 Submit Report
  Future<void> _submitReport() async {
    if (_isSubmitting) return;

    String title = titleController.text.trim();
    String description = descriptionController.text.trim();

    if (title.isEmpty || description.isEmpty || _image == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text("Please fill all fields")),
      );
      return;
    }

    setState(() => _isSubmitting = true);

    // Save report
    Report newReport = Report(
      title: title,
      description: description,
      imagePath: _image!.path,
    );

    ReportService.reports.add(newReport);

    // Show success message
    ScaffoldMessenger.of(context).showSnackBar(
      const SnackBar(content: Text("Report Submitted Successfully")),
    );

    await Future.delayed(const Duration(seconds: 1));

    if (!mounted) return;

    setState(() => _isSubmitting = false);

    titleController.clear();
    descriptionController.clear();
    _image = null;

    Navigator.pop(context);
  }

  @override
  void dispose() {
    titleController.dispose();
    descriptionController.dispose();
    _speech.stop();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text("Report Issue"),
      ),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(16),
        child: Column(
          children: [
            // Title Field
            TextField(
              controller: titleController,
              decoration: const InputDecoration(
                labelText: "Issue Title",
                border: OutlineInputBorder(),
              ),
            ),

            const SizedBox(height: 15),

            // Description Field
            TextField(
              controller: descriptionController,
              maxLines: 4,
              decoration: const InputDecoration(
                labelText: "Description",
                border: OutlineInputBorder(),
              ),
            ),

            const SizedBox(height: 15),

            // Camera Button
            ElevatedButton(
              onPressed: _pickImage,
              child: const Text("Capture Image 📷"),
            ),

            const SizedBox(height: 10),

            // Image Preview
            _image != null
                ? ClipRRect(
              borderRadius: BorderRadius.circular(10),
              child: Image.file(
                _image!,
                height: 150,
                width: double.infinity,
                fit: BoxFit.cover,
              ),
            )
                : const Text("No image selected"),

            const SizedBox(height: 15),

            // Microphone Button
            ElevatedButton(
              onPressed: _isListening ? _stopListening : _startListening,
              child: Text(
                _isListening
                    ? "Stop Recording 🎤"
                    : "Record Voice 🎤",
              ),
            ),

            const SizedBox(height: 20),

            // Submit Button
            ElevatedButton(
              onPressed: _isSubmitting ? null : _submitReport,
              child: _isSubmitting
                  ? const SizedBox(
                height: 20,
                width: 20,
                child: CircularProgressIndicator(strokeWidth: 2),
              )
                  : const Text("Submit Report"),
            ),
          ],
        ),
      ),
    );
  }
}