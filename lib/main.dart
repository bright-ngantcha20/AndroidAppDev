import 'package:flutter/material.dart';
import 'screens/login_screen.dart';
import 'screens/home_screen.dart';
import 'screens/report_screen.dart';
import 'screens/reports_list.dart';
import 'screens/report_details_screen.dart';

void main() {
  runApp(const MyApp()); // ✅ add const
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Campus Report',
      debugShowCheckedModeBanner: false,

      // ✅ Optional theme (nice upgrade)
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),

      initialRoute: '/',

      routes: {
        '/': (context) => const LoginScreen(),
        '/home': (context) => const HomeScreen(),
        '/report': (context) => const ReportScreen(),
        '/reports': (context) => const ReportListScreen(),
        '/details': (context) => const ReportDetailsScreen(),
      },
    );
  }
}