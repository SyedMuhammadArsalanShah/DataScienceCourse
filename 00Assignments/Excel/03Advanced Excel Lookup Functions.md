
### **Assignment Title: Comprehensive Marksheet, Report Card & Quran Data Card Using Advanced Excel Lookup Functions**

---

### **Objective:**
The objective of this assignment is to create a functional **Student Marksheet and Report Card** and a **Quran Data Card** using Excel’s lookup and data analysis functions such as `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `SUM`, `SUMIF`, `COUNTIF`, and more. This assignment will test your ability to effectively apply these functions to real-world data sets.

---

### **Part 1: Student Marksheet and Dynamic Report Card**

**Scenario:**
You are provided with a dataset of students, including their Student IDs, names, and subject marks. Your task is to generate a dynamic and automated marksheet and report card using advanced Excel functions.

---

#### **Tasks & Instructions:**

1. **Dataset Setup:**
   - Create a table that includes:
     - **Student ID** (Unique identifier)
     - **Student Name**
     - **Subjects** (Math, English, Science, etc.)
     - **Marks obtained** in each subject

2. **Marksheet Creation:**
   - **VLOOKUP**: Use `VLOOKUP` to dynamically fetch the marks for each student by entering their **Student ID**.
   - **HLOOKUP**: Use `HLOOKUP` to retrieve marks or subject-related data horizontally.
   - **SUM**: Calculate the **Total Marks** for each student by summing up their marks in all subjects using the `SUM` function.
   - **Percentage Calculation**: Calculate the percentage by dividing the **Total Marks** by the total possible marks for all subjects.
   - **Grade Assignment**: Use `IF` statements or a grading table (via `VLOOKUP` or `XLOOKUP`) to assign grades based on the percentage (e.g., A, B, C, D, F).

3. **Report Card Analysis:**
   - **SUMIF**: Calculate the total marks for students who scored above a certain threshold in a specific subject, for example, above 80 in Math.
   - **SUMIFS**: Use `SUMIFS` to calculate the total marks of students who scored above a specific threshold in two or more subjects (e.g., above 80 in both Math and Science).
   - **COUNT**: Use the `COUNT` function to count the number of subjects for each student.
   - **COUNTA**: Use `COUNTA` to count the number of students in the sheet.
   - **COUNTIF**: Use `COUNTIF` to count how many students scored above a certain mark in a subject (e.g., above 70% in Math).
   - **COUNTIFS**: Use `COUNTIFS` to count students who scored above a specific threshold in multiple subjects (e.g., more than 80 in both Math and English).
   - **XLOOKUP**: Implement `XLOOKUP` to create a dynamic lookup system for retrieving a student's marks and grade by entering their **Student ID** or **Name**.

4. **Bonus Task**:
   - Apply **Conditional Formatting** to highlight:
     - Marks above 80 in green.
     - Marks below 50 in red.
   - Create a **Drop-down list** that allows users to select a student's name, and automatically update the report card details based on the selected name.

---

### **Part 2: Quran Data Card Using Lookup Functions**

**Scenario:**
You are provided with a dataset of Surahs from the Quran, containing fields like Surah number, name, the number of verses, revelation type, and translation information. Your task is to create a dynamic Quran Data Card using Excel's lookup functions.

---

#### **Tasks & Instructions:**

1. **Data Preparation:**
   - Use the following fields in your dataset:
     - **Surah Number**
     - **Surah Name**
     - **Number of Verses (Ayat)**
     - **Revelation Type** (Makki/Madani)
     - **Translation Language**
     - **Translator’s Name**
     - **Recitation Audio URL** (optional)

2. **Quran Data Card:**
   - **XLOOKUP**: Create a dynamic Quran Data Card that retrieves details such as the Surah name, number of verses, and revelation type using `XLOOKUP`. This allows the user to input the Surah number and get its corresponding information.
   - **VLOOKUP**: Alternatively, use `VLOOKUP` to return similar information (Surah name, number of verses, revelation type) by entering the Surah number.
   - **HLOOKUP**: Use `HLOOKUP` to search horizontally for Surah data (e.g., number of Ayat, revelation type) across a row of Surah numbers.

3. **Data Analysis:**
   - **COUNTIF**: Count the number of Makki and Madani Surahs using `COUNTIF`.
   - **SUMIF**: Use `SUMIF` to calculate the total number of verses in all **Makki** Surahs.
   - **SUMIFS**: Use `SUMIFS` to calculate the total number of verses in **Madani** Surahs that have more than a certain number of verses (e.g., more than 50 verses).
   - **LOOKUP**: Use `LOOKUP` to retrieve the translator’s name based on the translation language for any specific Surah.

4. **Bonus Task**:
   - Implement **Hyperlinks** for the recitation URLs so that clicking on a Surah name opens the online recitation of that Surah.
   - Add a **Drop-down list** to dynamically select a Surah, and update the Data Card automatically with its information.

---

### **Submission Requirements:**

1. **Excel File**: Submit the Excel workbook with the following sheets:
   - **Sheet 1**: **Student Marksheet and Report Card**, implementing all the required functions and formulas.
   - **Sheet 2**: **Quran Data Card**, showing how the lookup functions have been applied to retrieve Surah information dynamically.

2. **Documentation**: Provide a brief explanation (in a separate document or as notes in the Excel file) describing how you applied each function and formula in the assignment.

3. **Submission Deadline**: Please submit your completed assignment by **October 10, 2024**, via **GitHub**. Ensure that your repository is publicly accessible or shareable.

---

### **Grading Criteria:**

- **Accuracy of Lookup Functions**: (30%)
  - Correct use and application of `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, and `LOOKUP`.
- **Use of Aggregate Functions**: (30%)
  - Proper use of `SUM`, `SUMIF`, `COUNT`, `COUNTIF`, and related functions.
- **Presentation & Formatting**: (20%)
  - Neat and clear formatting, proper organization, and use of conditional formatting where required.
- **Bonus Features & Creativity**: (20%)
  - Extra features such as conditional formatting, drop-down lists, hyperlinks, etc.

---

**Good luck!** This assignment will help you enhance your understanding of advanced Excel functions and how they can be applied to real-world scenarios.

