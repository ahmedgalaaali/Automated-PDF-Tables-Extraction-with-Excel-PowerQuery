# Automated-PDF-Tables-Extraction-with-Excel-PowerQuery
## Content
1. [Overview](#overview)
2. [Extraction steps](#extraction-steps)
3. [Transformation Example](#transformation-example)
4. [Where's the Automation?](#wheres-the-automation)
5. [Conclusion](#conclusion)
## Overview
Excel Power Query offers amazing options that eliminate the need for repetitive tasks, which can be overwhelming when dealing with extensive data transformations. In this example, I will create a folder connected to Excel Power Query. This folder will extract tables containing employee data, transform the data, and automatically apply the same transformation steps each time a new `pdf` is added to the folder.
## Extraction steps
1. Create a folder that contains the `pdf` files from which the tables to be exctracted.
2. From **Data** tab in excel, choose **Get data** > **From file** > **From folder**
3. Choose the created folder containing the `pdf`, the connection window of PowerQuery will launch showing all the files contained in the folder. Choose **Combine & Transform**.
   <div align="center">
   <img src="https://github.com/user-attachments/assets/557f3d07-d4ba-4199-8a55-6dad5968a632" alt="Screenshot 2025-01-21 161417" width="600px">
   </div>
4. From the **Combine Files** window, choose a **Sample Files** and below click on **Parameter1** and then **OK**.
   <div align="center">
   <img src="https://github.com/user-attachments/assets/2e794f47-6f91-4004-a762-4bca456f957e" alt="pic">
   </div>
6. This will launch PowerQuery editor to begin transforming the tables, filter the **Kind** column to choose only the **Tables** not **Pages**.
   <div align="center">
   <img src="https://github.com/user-attachments/assets/762c90e5-4a24-476b-844f-bbdff0a34c06" alt="pic">
   </div>
8. Click the **Expand** button at the top of the **Data** column, this will expand the existing tables and show them in the PowerQuery edetor, confirm the expand process.
   <div align="center">
   <img src="https://github.com/user-attachments/assets/63b06c51-d35f-4e50-9e2c-bd037450869b" alt="pic">
   </div>
10. **Remove the first 4 columns** that contain the table information after expanding, and **use the first row as headers**.  
11. **IMPORTANT** | Filter over one column of the table columns. You will find that it still contains the table name. **Why?** This happens because we expanded **multiple tables** that existed in the file. To easily get rid of them, uncheck the name from the filter button.  

    <div align="center">
      <img src="https://github.com/user-attachments/assets/5b69f6b1-82ac-49f0-913e-247d0342d990" alt="Screenshot 2025-01-21 161746">
    </div>  

12. These steps are more than enough to **close & load** the tables into Excel. When loaded, this shows all the rows of [the first `pdf`](https://github.com/ahmedgalaaali/Automated-PDF-Tables-Extraction-with-Excel-PowerQuery/blob/142f164514d40c7c5c19c0498d2752829be1d6f7/Employee%20Remuneration.pdf) are loaded to the Excel file as shown below.  

    <div align="center">
      <img src="https://github.com/user-attachments/assets/4e7e7f07-358c-49f4-a68f-b3c9482a113e" alt="Screenshot 2025-01-21 161838">
    </div>  

## Transformation Example  
Suppose that your manager assigned you a task to be done to the sent `pdf` files. He wants to extract only the `JOB DESCRIPTION` and get rid of the `DEPARTMENT` from the `DESIGNATION` column. _(All the details are in this file [here](https://github.com/ahmedgalaaali/Automated-PDF-Tables-Extraction-with-Excel-PowerQuery/blob/142f164514d40c7c5c19c0498d2752829be1d6f7/Employee%20Remuneration.pdf))_. Alright, I got you:  

1. Open the PowerQuery editor again and navigate to the `DESIGNATION` column. Right-click on it, choose **Split Column** > **By Delimiter**. In our case, the delimiter is ` (`. Type it in the dialog box and click **OK**.  

    <div align="center">
      <img src="https://github.com/user-attachments/assets/272ba6a6-320f-4a12-8963-d6a8c0d0c1f3" alt="Screenshot 2025-01-21 161936">
    </div>  

2. This will result in two columns. Rename the first one to `JOB DESCRIPTION` and remove the second one.  

    <div align="center">
      <img src="https://github.com/user-attachments/assets/652e7aa3-1907-4506-865e-1c024edcc814" alt="Screenshot 2025-01-21 162036">
    </div>  

3. With that, we’ve completed the task. _And here's my favorite part_: **Close & Load**.  

## Where's the Automation?  
Well, that's a good question. PowerQuery keeps all the transformation steps so that it can apply them to additional files in the future. Here's a real-world scenario:  

1. Your manager sends you a new version of the previously sent file. It has the same structure and needs to be transformed the same way and added to the existing data.  
2. Simply, place the file in the chosen folder from the first step.  
3. In Excel, refresh the query, and you'll see the new data added to the query.  

    <div align="center">
      <img src="https://github.com/user-attachments/assets/677337a9-60e0-483f-8367-24ab928216ba" alt="Screenshot 2025-01-21 162147">
    </div>  

## Conclusion  
Finally, we accomplished two main tasks:  
1. Extracted data from the `pdf` file.  
2. Linked a folder to automate the same steps for future files.  

This ensures we avoid repetitive tasks, especially when dealing with a large number of files and pages.  

