![Screenshot 2025-01-21 162147](https://github.com/user-attachments/assets/e4df1048-377b-4e7e-8092-0c868ea1c2c9)# Automated-PDF-Tables-Extraction-with-Excel-PowerQuery
## Overview
Excel Power Query offers amazing options that eliminate the need for repetitive tasks, which can be overwhelming when dealing with extensive data transformations. In this example, I will create a folder connected to Excel Power Query. This folder will extract tables containing employee data, transform the data, and automatically apply the same transformation steps each time a new `pdf` is added to the folder.
## Extraction steps
1. Create a folder that contains the `pdf` files from which the tables to be exctracted.
2. From **Data** tab in excel, choose **Get data** > **From file** > **From folder**
3. Choose the created folder containing the `pdf`, the connection window of PowerQuery will launch showing all the files contained in the folder. Choose **Combine & Transform**.
   ![Screenshot 2025-01-21 161417](https://github.com/user-attachments/assets/557f3d07-d4ba-4199-8a55-6dad5968a632)
4. From the **Combine Files** window, choose a **Sample Files** and below click on **Parameter1** and then **OK**.
   ![Screenshot 2025-01-21 161442](https://github.com/user-attachments/assets/2e794f47-6f91-4004-a762-4bca456f957e)
5. This will launch PowerQuery editor to begin transforming the tables, filter the **Kind** column to choose only the **Tables** not **Pages**.
   ![Screenshot 2025-01-21 161547](https://github.com/user-attachments/assets/762c90e5-4a24-476b-844f-bbdff0a34c06)
6. Click the **Expand** button at the top of the **Data** column, this will expand the existing tables and show them in the PowerQuery edetor, confirm the expand process.
  ![Screenshot 2025-01-21 161617](https://github.com/user-attachments/assets/63b06c51-d35f-4e50-9e2c-bd037450869b)
7. **Remove the first 4 columns** that contain the tables info after expanding, **use the first row as headers**.
8. **IMPORTANT** | Filter over on column of the table columns, you will find that it still contains the table name. **Why?** Simply, because we expanded **multiple tables** that existed at the file. To easily get rid of them, uncheck the name from the filter button.
   ![Screenshot 2025-01-21 161746](https://github.com/user-attachments/assets/5b69f6b1-82ac-49f0-913e-247d0342d990)
9. **These steps are more than enough to be ready to **close & load** the tables into excel. When loaded, this shows all the rows of [the first `pdf`](https://github.com/ahmedgalaaali/Automated-PDF-Tables-Extraction-with-Excel-PowerQuery/blob/142f164514d40c7c5c19c0498d2752829be1d6f7/Employee%20Remuneration.pdf) are loaded to the excel file as shown below.
    ![Screenshot 2025-01-21 161838](https://github.com/user-attachments/assets/4e7e7f07-358c-49f4-a68f-b3c9482a113e)
## Transformation Example
Suppose that your manager assigned you a task to be done to the sent `pdf` files, he want to extract only the `JOB DESCRIPTION` and get rid of the `DEPARTMENT` from the `DESIGNATION` column. _(All the details are in this file [here](https://github.com/ahmedgalaaali/Automated-PDF-Tables-Extraction-with-Excel-PowerQuery/blob/142f164514d40c7c5c19c0498d2752829be1d6f7/Employee%20Remuneration.pdf))_. Alright I got you:
1. Open the PowerQuery editor agian and navigate to the `DESIGNATION` column, right click on it, choose **spit column** > **by delimeter**. In our case, the delimeter is ` (`. Type it at the dialoge box and click **OK**
   ![Screenshot 2025-01-21 161936](https://github.com/user-attachments/assets/272ba6a6-320f-4a12-8963-d6a8c0d0c1f3)
2. This will end up with 2 columns, change the name of the first one to `JOB DESCRIPTION` and remove the second one.
   ![Screenshot 2025-01-21 162036](https://github.com/user-attachments/assets/652e7aa3-1907-4506-865e-1c024edcc814)
3. With that, we performed the task, _and here's my favourite part_: **Close & Load**.
## Where's the automation?
Well, that's a good question. PowerQuery keeps all the transformation steps so that it can apply it to more files to be added in the future, here's a realworld scenario:
1. Your manager sends you a new verion of the file sent before, it has the same structure, so that it need to be transformed the same way and added to the existing data.
2. Simply, take the file and put it in the choosen folder of the very first step.
3. In exce, refresh the query and you will notice adding the new data to the query.
   ![Screenshot 2025-01-21 162147](https://github.com/user-attachments/assets/677337a9-60e0-483f-8367-24ab928216ba)
## Conclusion
Finally, we performed 2 main tasks:
1. Extracted the data from `pdf` file.
2. Linked a folder to automate the same steps again when needed to.
This ensures not doing repetitive tasks so many times especially when the files are in greate numbers and different pages.
