# Decision support system used to build priority queues of game strategies
The system is based on Wald-Savage synthetic criterion used in "games against nature" (game theory).
This system was used by me to solve the problem of building the priority queue of the commercial bank mortgage loan recipients.

The program collects information by connecting to a relational database (such as MySQL) stored on a remote server, and reading information from it using SQL queries, which is provided by the MySQL Connector/NET 1.3.8 library. Information is systematized within the program by entering it into a specialized data structure - a 2-dimensional array.
The Wald-Savage algorithm is formalized using C# with Windows Forms GUI on a .NET Core 5.0.
After the information has been processed by the algorithm, it is displayed in the control element of Windows Forms GUI - DataGridView. Note: As a data source that can be defined by a DataSource property of a DataGridView element, this element can use various types of data storage (including database tables) or work without binding to any data source. The optional function is to output the processed to the MS Excel file, which is implemented with the Open-XML-SDK 2.9.1 package. This option can be applied if the user puts a check box next to the inscription «Output to MS Excel».

Screenshot below shows the program window at startup:
![image](https://user-images.githubusercontent.com/76261338/189528327-dc088e28-7b38-48cd-8f1f-fd0cb7066da6.png)

After filling in all the corresponding fields, press the button shown in the screenshot below to connect to the database using the input data:
![image](https://user-images.githubusercontent.com/76261338/189528446-17fdc2ba-515d-4d37-b8d4-1ff37b304ca0.png)

After connecting to the database, the program window will look like the next screenshot:
![image](https://user-images.githubusercontent.com/76261338/189528464-9a76b152-0c30-45dd-b3c4-cda241de2c36.png)

After the necessary database has been connected and data has been unloaded, to form priority queues of game strategies it is necessary to press the button shown in the screenshot:
![image](https://user-images.githubusercontent.com/76261338/189528529-7dbbcc7b-d87b-443f-978d-dcc98bd838ac.png)

After clicking the above button, the algorithm will be launched to build a set of strategies optimally according to the Wald-Savage criterion, and then an algorithm to build a set of priority queues of game strategies with all the necessary results displayed:
![image](https://user-images.githubusercontent.com/76261338/189528630-d94a618a-5b76-4d2b-8f13-24b60508fdab.png)

If it is necessary that the set of priority queues of game strategies formed is written into a file, before pressing the button responsible for running the algorithm, you need to put a check box opposite the inscription «Output to MS Excel file»:
![image](https://user-images.githubusercontent.com/76261338/189528785-9efc7c4c-5ae6-46e7-8413-0c888bef1c91.png)

An examle of the sheet in the MS Excel workbook, in which a number of priority queues of game strategies were written:
![image](https://user-images.githubusercontent.com/76261338/189528789-170f31a3-a054-46ef-8804-b2e80aa2062f.png)

For a deeper understanding I advise you to read: Labsker, L. G. Priority of lending to the bank of corporate borrowers. Formation of the priority order on the basis of the synthetic criterion of Wald-Savage / L. G. Labsker, N. A. Yashchenko, A. V. Ameline. - Saarbrüken: LAP LAMBERT Academic Publishing GmbH & Co. KG, 2012. - 237 p.
