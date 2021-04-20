# Sustainability-rep
Before using the program, read this and the file named IPR.pdf. 
This README provides instructions on how to use this program as a Google doc add-on 

<b> NOTE: steps 1 to 3 are only needed when the add-on is first taken into use, only one persons from the people working on the file needs to perform them. Instructions for starting the project when the script is already attached is described in step 5 and is not a required step for the one who attaches the script to the file </b> 
 
<b> Step 1: Opening script editor in a file </b> <br>
  * 1.1. Open the google docs file in which you want to use the add-on <br>
  * 1.2. Click on "Tools" in the toolbar of the page <br>
  * 1.3. Click on "Script editor" <br>
  * 1.4. The script editor opens in a new tab of your browser <br>
  * 1.5. Rename the project by clicking on "Untitled project" at the top of the page and typing a name such as "Susaf-tool" 

<b> Step 2: Attaching the script </b> <br>
  * 2.1. Rename the "code.gs" file to "main.gs" (by clicking on the three dots that appear when hovering over the name and clicking on "rename") <br>
  * 2.2. Copy the entire script from "main.gs" of this repository <br>
  * 2.3. Go back to the "main.gs" file that you just renamed and paste the copied script into the "main.gs" content (the existing text in the file must be erased) <br>
  * 2.4. Create the sidebar.html file <br> 
  		* 2.4.1. Next to "Files" click on the "+" <br> 
       * 2.4.2. Click on HTML <br>
       * 2.4.3. Enter "sidebar" in the field <br>
       * 2.4.4. Erase the existing text in the file <br> 
       * 2.4.5. Copy the script from "sidebar.html" from this repository and paste it in the file of the same name in Apps Script   
  * 2.5. Create the popup.html file <br> 
       * 2.4.1. Next to "Files" click on the "+" <br> 
       * 2.4.2. Click on HTML <br>
       * 2.4.3. Enter "popup" in the field <br>
       * 2.4.4. Erase the existing text in the file <br> 
       * 2.4.5. Copy the script from "popup.html" from this repository and paste it in the file of the same name in Apps Script
  * 2.5. Create the categories.html file <br> 
       * 2.4.1. Next to "Files" click on the "+" <br> 
       * 2.4.2. Click on HTML <br>
       * 2.4.3. Enter "categories" in the field <br>
       * 2.4.4. Erase the existing text in the file <br> 
       * 2.4.5. Copy the script from "categories.html" from this repository and paste it in the file of the same name in Apps Script
	
<b> Step 3: Getting the add-on visible in google docs </b> <br>
* 3.1. Go to the main.gs file and click on "run" <br>
* 3.2. In the popup "Authorization required" click "review permissions" <br> 
* 3.3. Choose the account with which you want to use the add-on <br>
* 3.4. In the window "google hasn't verified this app" click "advanced" and "go to <the name of the project> (unsafe)" <br>
* 3.5. In the new window you will see what is required by the app (accessing to drive, in order to see if you have authorization to edit in a folder and to create a new file in this folder, edit google docs files, access to spreasheet in order to create one and display things in the user-interface such as the sidebar), click "Allow" <br>
* 3.6. Go back to the google docs file and reload the page (close the prompt if there is one), now the script is attached to the file. <br> 
	
<b> Step 4: Getting the script running </b> <br>
  * 4.1. In the toolbar click on "Add-ons" <br>
  * 4.2. Hover over the <project name> (the displayed name will be the one you defined in 1.5.)
  * 4.3. Click on "Start" <br>
  * 4.4. The sidebar gets visible next to the google docs file <br>
	
	
<b> Step 5: Getting the script running when it is already attached to the file </b> </br> 
  * 5.1. In the toolbar of the google docs file click on "Add-ons" <br>
  * 5.2. Hover over the <project name> <br>
  * 5.3. Click on "Start" <br>
  * 5.4. A prompt appears to ask for you permission, click on "continue" <br>
  * 5.5. Select the account on which you want to use the file <br>
  * 5.6. Click on "advanced" and on "go to <the name of the project> (unsafe)" <br>
  * 5.7. Click on "allow" and you now have authorized the script to work 
  * 5.8. Start the project again by going to add-on -> hovering over project name -> clicking "start"
  * 5.9. The sidebar shows on the side of the opened google docs file 

	  
