# Calendar Data Export Tool
A tool for generating lunar/solar calendar events and exporting them to Excel format.

## Features
- Generate lunar/solar calendar events for specified date range
- Include solar terms and auspicious/inauspicious activities
- Include public holidays and work arrangements
- Export data to Excel format
- Support custom file save path

## Installation
### Clone the repository
```bash
git clone git@github.com:XuQuan-nikkkki/calendar_data_export_tool.git
```

### Enter project directory
```bash
cd calendar_data_export_tool 
```

### Install dependencies
```bash
npm install
```

### Run the program
```bash
node index.js
```

Follow the prompts to input:
1. Start date (format: YYYY-MM-DD)
2. End date (format: YYYY-MM-DD)

### Note
- The program will generate a file named `Calendar_Data_${startDate}_${endDate}.xlsx` in the `Downloads` folder of your home directory.

