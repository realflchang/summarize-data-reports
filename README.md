<h1>Summarize reports.xlsm - Prerequisites and Instructions</h1>
This macro is to inspect several Excel data reports, and then provide a simple summary of the data reports.  The summary is in its own worksheet.  Useful when delivering to another staff and the staff does not need to open each individual Excel data reports, he or she can just see the summary worksheet.<br />

<!---
<h2>Video Demonstration</h2>

- ### [YouTube: How To Install osTicket with Prerequisites](https://www.youtube.com) -->

<h2>Environments and Technologies Used</h2>

- Microsoft Excel
- Visual Basic / Macros

<h2>Operating Systems Used </h2>

- Windows 10</b> (21H2)

<h2>List of Prerequisites</h2>

- Computer running Windows 10. Possible to run in Mac but I have not tested it.
- Excel summarize data report macro file – contains the VB script (“summarize reports.xlsm”)
- Macros in Excel are enabled
- Other individual Excel data report files

<h2>Steps</h2>

<p>1.	Create a folder in Windows, My Documents folder to store Excel macros. Ie. My Documents\Macros</p>

<p>2.	Download the Excel macro file (“reformat data report.xlsm”) into that folder</p>

<img src="https://github.com/user-attachments/assets/688a58df-7b83-4f4d-baa2-070da01615fd" alt="Download the summarize report.xlsm file into My Documents Macros folder" />

<p>3.	Next, we want to allow Excel to be able to run the Macro. There are 2 ways:</p>

  *	One way is to click “Enable Content” when this .xlsm file is opened. Problem with this is this will have to be done every time the macro file is opened.
<img src="https://github.com/user-attachments/assets/be5559b0-f82c-4645-9faf-e12d8fe7043d" alt="Macros have been disabled Security Warning" />

  *	Another more permanent way is to allow Excel to trust the location of the .xlsm file. 
    -	To do that, go to File, then click “Options” at the bottom of the menu
<img src="https://github.com/user-attachments/assets/8c7df4c9-9649-4f5a-8de4-fe0b34f8ce2e" alt="Menu File Options" />

    - Click on Trust Center, Trust Center Settings…
<img src="https://github.com/user-attachments/assets/33e1aa94-9133-4b17-bf6a-d6b60a9fb7d1" alt="Trust Center Settings" />

    - Trusted Locations, Add new location…
<img src="https://github.com/user-attachments/assets/c190314e-e4ef-43fe-a41e-6074e4771f91" alt="Add new location" />

    - Browse to the Macros folder that was just created and click OK:
<img src="https://github.com/user-attachments/assets/79751c53-4235-4f35-b18a-997a204e6755" alt="Browse to Macros folder" />

    - Added new location:
<img src="https://github.com/user-attachments/assets/b0ce28c0-f958-4150-b9c7-ad1c09c3292e" alt="Added new location" />

    - Next time we open the .xlsm file in Macros folder, there will not be a tooltip message from Excel 

<p>4.	Next, we need to provide a summary of individual data reports to our other department staff.  In my case, I have already prepared 4 individual data reports:</p>

<img src="https://github.com/user-attachments/assets/bde320f7-27cb-4e3e-860c-b21b2236bf6d" alt="Individual report 1" /><br />
<img src="https://github.com/user-attachments/assets/5ac2be85-9832-433a-863e-61d492d75e02" alt="Individual report 2" /><br />
<img src="https://github.com/user-attachments/assets/ba1e98e0-157a-46a4-878d-af521ae83133" alt="Individual report 3" /><br />
<img src="https://github.com/user-attachments/assets/6b9ec5ef-9db9-445f-94ed-ad56c8aa7996" alt="Individual report 4" /><br />

<p>5.	I have these 4 individual data reports saved:</p>

<img src="https://github.com/user-attachments/assets/f1e85c5d-8ef1-4d1a-b83f-ec6ee3cfeac6" alt="Saved individual reports" /><br />

<p>6.	Next, I open my summarize reports.xlsm:</p>

<img src="https://github.com/user-attachments/assets/3fa76b21-bbf5-4c63-8039-2c9999379bd6" alt="Opened .xlsm file" /><br />

<p>7.	While I have the individual data reports still open (if closed, please just reopen them), I drag each worksheet into the summary reports.xlsm so they are all in the same worksheet:</p>

<img src="https://github.com/user-attachments/assets/c78854f3-0789-49ff-81be-2852b5a4980b" alt="Sheet 1" /><br />
<img src="https://github.com/user-attachments/assets/0c7cd9bf-dd2c-4748-8a05-f6c3ee134beb" alt="Sheet 2" /><br />
<img src="https://github.com/user-attachments/assets/bea22329-a089-4854-8529-5e5a555b8bf0" alt="Sheet 3" /><br />
<img src="https://github.com/user-attachments/assets/e90efae2-e662-4573-a815-2fb7feef95d7" alt="Sheet 4" /><br />

<p>8.	To run the macro, go to View, Macros, View Macros…:</p>

<img src="https://github.com/user-attachments/assets/b5d73dc2-a629-4f7f-93fc-cc2e3bbcd77f" alt="View Macros" />

<p>9.	Click on the Macro from the .xlsm file that includes the name “RenameSheets”, and click Run:</p>
<img src="https://github.com/user-attachments/assets/51f6f91d-d39f-47f7-a4b4-655112024ab7" alt="Run Macro" />

  * The Options… button provides an option to assign a key shortcut, so that in the future, instead of browsing to View, Macros, View Macros, we can just do the key shortcut combination to run the same macro.

<img src="https://github.com/user-attachments/assets/60dbe462-4203-47f4-a54f-d6412872ad85" alt="Run Macro option key shortcut" />

<p>10.	Result of the macro. Result is the 4 individual worksheets are summarized in the “Sum” worksheet, in my case I have the Code and Balance of each individual worksheet, and the total of the balances:</p>

<img src="https://github.com/user-attachments/assets/347ed4ad-f169-4c97-98bb-27113e42b8f2" alt="Result" />

<p>11.	This macro can be customized to perform other formats as well.  With this, our colleagues can perform other tasks that would normally be manual and may result in inaccuracies and inefficient work.</p>



