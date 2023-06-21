# SplitExcelFile

This mini-project is used to split or join excel files.
Maybe I'll expand it to run other batches.
It's highly customized for my needs.
Feel free to copy it and edit it.

### Reminder!!!

> Always open the excel file and click on Enable Editing

### To Run the project

If you wanna split a big file, set the .env config NUMBER_ROWS_OUTPUT
Put the big excel file in the folder **excel_to_split** (create it it you didnt yet) in the root of your repository.
Then run

> npm run start-split

If you want to join files a single excel, you need to put the files in the **excel_to_join** folder.
Then run

> npm run start-join

### Order of execution:

(this is just to remember how the files are running)

1. server.js
2. app.js
3. express.js
4. splitExcel.js or joinExcel.js

### Configuring the folders:

See the .env.example
Create the folder in the root of the repository
