const fs = require('fs');
const XLSX = require('xlsx');

function printReport(pages) {
    console.log("===========")
    console.log("REPORT")
    console.log("===========")
    const sortedPages = sortPages(pages)
    for (const sortedPage of sortedPages) {
        const url = sortedPage[0]
        const hits = sortedPage[1]
        console.log(`Found ${hits} links to page: ${url}`)
    }
    console.log("===========")
    console.log("END REPORT")
    console.log("===========")
    
    const reportData = generateReportData(sortedPages);
    downloadReport(reportData);
}

function sortPages(pages) {
    const pagesArr = Object.entries(pages)
    pagesArr.sort((a, b) => {
        aHits = a[1]
        bHits = b[1]
        return b[1] - a[1]
    })
    return pagesArr
}

function generateReportData(sortedPages) {
    let reportData = "===========" + "\n";
    reportData += "REPORT" + "\n";
    reportData += "===========" + "\n";
    
    for (const sortedPage of sortedPages) {
        const url = sortedPage[0];
        const hits = sortedPage[1];
        reportData += `Found ${hits} links to page: ${url}` + "\n";
    }
    
    reportData += "===========" + "\n";
    reportData += "END REPORT" + "\n";
    reportData += "===========" + "\n";
    
    return reportData;
}

function downloadReport(reportData) {
    fs.writeFile('website_report.txt', reportData, (err) => {
        if (err) {
            console.error("Error occurred while downloading the report:", err);
        } else {
            console.log("Report downloaded successfully.");
            convertToExcel(reportData);
        }
    });
}

function convertToExcel(reportData) {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([[reportData]]);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Website Report");
    const excelFilePath = 'website_report.xlsx';
    XLSX.writeFile(workbook, excelFilePath);
    console.log("Excel report created:", excelFilePath);
}

module.exports = {
    sortPages,
    printReport,
    downloadReport
}