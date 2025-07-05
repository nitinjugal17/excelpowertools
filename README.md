
# Excel Power Tools

Excel Power Tools is a versatile, web-based utility designed to automate and streamline a wide range of common and complex tasks in Microsoft Excel. Built with modern web technologies, it processes files entirely on the client-side, ensuring your data remains private and secure. From splitting large files into manageable sheets to performing complex data aggregation and formatting, this application provides a powerful suite of tools for data analysts, administrators, and anyone who works extensively with spreadsheets.

![Excel Power Tools Screenshot](https://placehold.co/800x400.png?text=Excel+Power+Tools+UI)

## Table of Contents

- [About The Project](#about-the-project)
- [Tech Stack](#tech-stack)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
- [Features](#features)
  - [Data Organization & Structuring](#data-organization--structuring)
  - [Data Cleaning & Formatting](#data-cleaning--formatting)
  - [Analysis & Reporting](#analysis--reporting)
  - [AI-Powered Tools](#ai-powered-tools)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)
- [Frequently Asked Questions (FAQ)](#frequently-asked-questions-faq)

## About The Project

This application was developed to address the repetitive, time-consuming, and often error-prone tasks that Excel users face daily. Traditional methods often involve manual work or complex VBA macros. Excel Power Tools offers a user-friendly, browser-based alternative that is both powerful and secure.

**Core Principles:**

- **Privacy First**: All file processing happens in your browser. Your data is never uploaded to a server for processing, ensuring complete confidentiality.
- **User-Friendly Interface**: Complex operations are broken down into simple steps with clear instructions and options.
- **Efficiency**: Automate tasks that would take hours to complete manually, freeing you up to focus on analysis and decision-making.
- **Flexibility**: A wide range of tools and configuration options allows you to tailor each operation to your specific needs.

## Tech Stack

This project is built on a modern, robust, and scalable technology stack:

- **[Next.js](https://nextjs.org/)**: A React framework for building fast, server-rendered applications.
- **[React](https://react.dev/)**: A JavaScript library for building user interfaces.
- **[TypeScript](https://www.typescriptlang.org/)**: A statically typed superset of JavaScript that adds type safety.
- **[Tailwind CSS](https://tailwindcss.com/)**: A utility-first CSS framework for rapid UI development.
- **[shadcn/ui](https://ui.shadcn.com/)**: A collection of beautifully designed, reusable components.
- **[Genkit (for AI)](https://firebase.google.com/docs/genkit)**: A toolkit for building production-ready AI-powered features.
- **[xlsx-js-style](https://www.npmjs.com/package/xlsx-js-style)**: A library for reading, manipulating, and writing spreadsheet files with styling.

## Getting Started

To get a local copy up and running, follow these simple steps.

### Prerequisites

Ensure you have the following installed on your system:

- **Node.js** (v18.x or later)
- **npm** (v9.x or later) or **yarn**

### Installation

1.  **Clone the repository:**
    ```sh
    git clone https://github.com/your-username/excel-power-tools.git
    cd excel-power-tools
    ```

2.  **Install NPM packages:**
    ```sh
    npm install
    ```

3.  **Set up environment variables:**
    Create a `.env.local` file in the root of the project and add any necessary environment variables (e.g., for AI services).
    ```
    GOOGLE_API_KEY=your_google_api_key_here
    ```

4.  **Run the development server:**
    ```sh
    npm run dev
    ```
    Open [http://localhost:3000](http://localhost:3000) with your browser to see the result.

## Features

Excel Power Tools offers a comprehensive suite of utilities to handle various spreadsheet tasks.

### Data Organization & Structuring

- **Sheet Splitter**: Automatically splits a single master sheet into multiple new sheets based on the unique values in a specified column.
- **Sheet Merger & Combiner**: Merge sheets from a source workbook into a destination file, or combine data from multiple sheets in one workbook into a single master sheet.
- **Workbook Breaker**: Break a large workbook with many sheets into multiple smaller, more manageable Excel files based on user-defined groups.
- **Column Purger**: Quickly remove one or more columns from multiple sheets at once, preserving all other data and formatting.
- **Excel Comparator**: Compare two Excel files sheet by sheet based on a primary key, generating a detailed report of new, deleted, and modified rows.
- **Pivot Table Creator**: Generate a pivot table from your data by specifying row, column, and value fields, all without needing to open Excel.

### Data Cleaning & Formatting

- **Duplicate Finder**: Identify and mark duplicate rows across multiple sheets based on a composite key. Optionally highlight duplicates or update a status column.
- **Empty Cell Finder**: Scan sheets for empty cells within a specified range or across all columns. Optionally highlights empty cells and generates a report.
- **Text Formatter**: Find text matching specific criteria (including regular expressions) and apply a wide range of formatting, such as font style, color, and cell fill.

### Analysis & Reporting

- **Data Aggregator**: A powerful tool to count occurrences of keywords based on custom mapping rules. It can operate in a simple value-match mode or a more structured key-match mode for summarizing categorized data.
- **Unique Value Finder**: Quickly extract and count all unique values from specified columns or substrings in text files, providing a clean list for analysis.
- **Data Extractor**: Find all rows that match a specific value in a lookup column and extract data from specified return columns into a new report.

### AI-Powered Tools

- **AI Smart Fill (Imputer)**: Uses generative AI to intelligently fill in blank cells by analyzing the context of surrounding data. It can also perform rule-based imputation, such as filling data based on the most common value in a group of duplicates.

## Usage

The general workflow for most tools in the application is as follows:

1.  **Select a Tool**: Choose the desired tool from the sidebar menu.
2.  **Upload Your File(s)**: Use the upload interface to select the source Excel file(s).
3.  **Configure Options**: Set the parameters for the operation, such as selecting sheets, specifying columns, and defining rules.
4.  **Process**: Click the main action button (e.g., "Analyze", "Split", "Format") to start the processing. All processing is done locally in your browser.
5.  **Review & Download**: After processing, review the results and download the newly generated workbook or report.

## Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".

1.  **Fork the Project**
2.  **Create your Feature Branch** (`git checkout -b feature/AmazingFeature`)
3.  **Commit your Changes** (`git commit -m 'Add some AmazingFeature'`)
4.  **Push to the Branch** (`git push origin feature/AmazingFeature`)
5.  **Open a Pull Request**

## License

This project is distributed under the GNU General Public License v3.0. See `LICENSE` for more information.

## Frequently Asked Questions (FAQ)

**Q: Is my data secure? Where are my files uploaded?**
**A:** Your data is 100% secure. All file processing happens entirely within your web browser on your computer. Your Excel files are **never** uploaded to any server.

**Q: What file types are supported?**
**A:** The application is optimized for modern Excel files (`.xlsx`, `.xlsm`). While it may work with older `.xls` files, compatibility is not guaranteed.

**Q: Why do some tools generate `.xlsm` (Macro-Enabled) files?**
**A:** Some tools, like the Sheet Updater, can generate VBScript macros that perform complex formatting inside the workbook itself. The `.xlsm` format is required to embed and run these macros. It also ensures that if your original file contained macros, they are preserved in the output. If you don't need macros, you can often choose the standard `.xlsx` output format.

**Q: Are there any limitations on file size?**
**A:** Since processing is done in the browser, performance depends on your computer's memory (RAM) and processing power. Very large files (hundreds of megabytes or millions of rows) may cause the browser to become slow or unresponsive. It is recommended to use the tools on reasonably sized files.

**Q: Can I use this tool on a Mac?**
**A:** Yes. Since this is a web application, it works on any operating system with a modern web browser (Chrome, Firefox, Safari, Edge). However, the generated VBScript macros are specific to Microsoft Excel on Windows and will not run on Excel for Mac. All other features work perfectly on any platform.
