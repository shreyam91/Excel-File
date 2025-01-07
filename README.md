# Excel Data Segregation Tool

A Java-based tool for reading, processing, and segregating Excel data into separate files or sheets based on specific criteria (e.g., a particular column value like "Region"). The tool utilizes **Apache POI** to handle Excel files (`.xls` and `.xlsx` formats).

---

## Features
- **Read Excel Files**: Supports both `.xls` and `.xlsx` formats.
- **Data Segregation**: Segregates data by custom criteria (e.g., based on values in a specific column).
- **Output**: Creates new Excel files or sheets for each segregated group.
- **Customizable**: Modify the segregation logic to fit your needs (e.g., by "Region", "Department", etc.).

## Requirements

- **Java 8 or higher**.
- **Apache POI**: For reading and writing Excel files.
- **Maven or Gradle**: For dependency management.

## Installation

### 1. Clone the Repository
### 2. Setup dependencies in Maven
<dependencies>
    <!-- Apache POI for handling Excel files -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.2.3</version> <!-- Latest version -->
    </dependency>
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi</artifactId>
        <version>5.2.3</version> <!-- Latest version -->
    </dependency>
</dependencies>

### 3. Build the project
      mvn clean install

### Troubleshooting

"File format is not supported"
Ensure the file you are working with is a valid .xls or .xlsx format.

"NullPointerException"
This may occur if there are empty rows or cells. Make sure your Excel file is well-structured.

"Excel file is too large"
If the Excel file is large, consider breaking it into smaller parts, or optimize memory usage for better performance.

### License

This project is licensed under the MIT License - see the LICENSE file for details.

### Contributing

1. Fork the repository.
2. Clone your fork to your local machine.
3. Create a new branch (git checkout -b feature-branch).
4. Commit your changes (git commit -am 'Add new feature').
5. Push to your branch (git push origin feature-branch).
6. Create a new Pull Request.

### Acknowledgments

1. Apache POI provides a powerful library for handling Excel files in Java.
2. OpenJDK for the Java runtime environment.
3. open the java
