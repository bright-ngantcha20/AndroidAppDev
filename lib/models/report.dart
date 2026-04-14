class Report {
  final String title;
  final String description;
  final String imagePath;
  String status;

  Report({
    required this.title,
    required this.description,
    required this.imagePath,
    this.status = "Pending",
  });
}