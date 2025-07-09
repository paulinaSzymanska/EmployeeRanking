# 🧮 Employee Ranking Calculator

This Java program calculates employee rankings based on Excel timesheet data. It scans a given directory (and its subdirectories) containing `.xls` and `.xlsx` files, extracts working hours, and displays rankings according to various criteria.

---

## ✅ Requirements

- **Java 21**
- **Gradle** (used for building and running the project)
- Internet access (for downloading dependencies from Maven Central)

---

## 📦 Sample Data

Sample timesheet data is organized in year/month subdirectories. You can find an example dataset here:
- ~\EmployeeRanking\reporter-dane\reporter-dane\2012


---

## 🚀 How to Run

1. **Clone or download** this repository.

2. In the terminal, navigate to the project root and build the app:

```bash
./gradlew build
```

3. Run the application:

```bash
./gradlew run
```

4. When prompted, enter the full path to the directory containing the Excel files (e.g., C:\Users\Example\EmployeeRanking\reporter-dane).

5. Enter the Excel file name when prompted

## 📘 Usage Scenarios
The program performs the following:

- Prompts the user to enter the path to the directory containing timesheet files.

- Reads all .xls and .xlsx Excel files recursively from that folder and its subfolders.

- Displays human-readable rankings based on aggregated working hours:

### a) 👷‍♂️ Employee Ranking
Displays employees ranked by total hours worked across all projects and all time periods.
Example:
```dtd
1. Anna Kowalska – 178.5 hours
2. Jan Nowak – 166.0 hours
```


### b) 🗓️ Month Ranking
Displays months ranked by total hours worked across all employees and all projects.
Example:
```dtd
1. January 2015 – 143 hours
2. February 2014 – 137 hours
```

### c) 📅 Top 10 Workdays
Displays the 10 most active workdays (aggregated across all employees and projects).
Example:
```dtd
1. 5 September 2015 – 43 hours
2. 4 September 2015 – 42 hours
```

## 🔧 Technologies Used
- Java 21
- Apache POI (for reading .xls and .xlsx)
- Gradle (build system)
- Log4j or SLF4J can be added for proper logging (currently logs may show POI warnings)

## 📬 Author
Created by Paulina Szymańska
For AGH course project submission

