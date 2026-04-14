import 'package:flutter/material.dart';
import 'package:shared_preferences/shared_preferences.dart';

class LoginScreen extends StatefulWidget {
  const LoginScreen({super.key});

  @override
  State<LoginScreen> createState() => _LoginScreenState();
}

class _LoginScreenState extends State<LoginScreen> {
  bool isLogin = true;

  final TextEditingController emailController = TextEditingController();
  final TextEditingController passwordController = TextEditingController();

  String message = "";

  // 🔐 Register User
  Future<void> registerUser() async {
    final prefs = await SharedPreferences.getInstance();

    String email = emailController.text.trim();
    String password = passwordController.text.trim();

    await prefs.setString("email", email);
    await prefs.setString("password", password);

    if (!mounted) return;

    setState(() {
      message = "Registration successful! Please login.";
      isLogin = true;
    });
  }

  // 🔓 Login User
  Future<void> loginUser() async {
    final prefs = await SharedPreferences.getInstance();

    String savedEmail = prefs.getString("email") ?? "";
    String savedPassword = prefs.getString("password") ?? "";

    String email = emailController.text.trim();
    String password = passwordController.text.trim();

    if (email == savedEmail && password == savedPassword) {
      if (!mounted) return;

      setState(() {
        message = "Login successful!";
      });

      // Determine role
      String role = email == "admin@gmail.com" ? "admin" : "student";

      // Navigate safely
      Navigator.pushReplacementNamed(
        context,
        '/home',
        arguments: role,
      );
    } else {
      if (!mounted) return;

      setState(() {
        message = "Invalid email or password";
      });
    }
  }

  @override
  void dispose() {
    emailController.dispose();
    passwordController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text(isLogin ? "Login" : "Register"),
      ),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          children: [
            TextField(
              controller: emailController,
              decoration: const InputDecoration(labelText: "Email"),
            ),

            const SizedBox(height: 10),

            TextField(
              controller: passwordController,
              decoration: const InputDecoration(labelText: "Password"),
              obscureText: true,
            ),

            const SizedBox(height: 20),

            ElevatedButton(
              onPressed: isLogin ? loginUser : registerUser,
              child: Text(isLogin ? "Login" : "Register"),
            ),

            TextButton(
              onPressed: () {
                setState(() {
                  isLogin = !isLogin;
                  message = "";
                });
              },
              child: Text(
                isLogin
                    ? "Don't have an account? Register"
                    : "Already have an account? Login",
              ),
            ),

            const SizedBox(height: 20),

            Text(
              message,
              style: const TextStyle(color: Colors.red),
            ),
          ],
        ),
      ),
    );
  }
}