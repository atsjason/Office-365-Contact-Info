### Future Changes
- [ ] Make the process alot faster when executing 
- [ ] Convert getters and setters to only two seperate object as calls can simply be made by passsing the object (Prevents API calls)
- [ ] Export CSV as attachment
 
 
 
 
# Office-365-Contact-Info

### This app will allow an Office 365 Administrator to update the Job Title, Department, and Supervisor fiedlds, all from a CSV file
##### Please make sure the formatting of the CSV file is maintained as such: 

![image](https://user-images.githubusercontent.com/98031074/155772716-a1c231ba-36ad-429a-97ba-61998561bb99.png)

The order of the columns do not matter; however, the names of the column itself do matter: 

**Last Name,First Name,Employee Email,Department,Job Title,Supervisor Name ,Supervisor Email Address**




### Location of the script
 ![image](https://user-images.githubusercontent.com/98031074/155773215-89a50cdf-3b85-4435-943d-e3487e2998ed.png)
##### This script will execute and will create a shortcut on the desktop for your convenience:
![image](https://user-images.githubusercontent.com/98031074/155773659-1bd40a61-e33d-4dca-a44d-d8f6f6ff0c60.png)


(Done by *Create_Shortcut.ps1*)


## Running script
##### This WILL require you to run the application as an adminstrator in order to successfully install the AzureAD connect module
##### This script will initially request for you to install AzureAD
##### **I would recommend to install the AzureAD module first, as the *Send email* portion of the app will require you to be on the correct VPN network**
###### **You only need to be on the VPN if you plan on emailing the results. the results will be provided via CSV and GUI form**

```
Install-module -name AzureAD
Import-module -name AzureAD
```

#### You will encounter a pop-up message when runnning each time to make sure you are infact ready to make changes
![image](https://user-images.githubusercontent.com/98031074/155775058-5e15d7f5-d9f1-4299-ad48-31585b4528e0.png)

#### Please be careful when making changes to Office 365
#### You will be prompted to access Office 365 admin each time
![image](https://user-images.githubusercontent.com/98031074/155775178-af9ba6bc-43e5-414c-a9c8-fcca308c7fa7.png)

You will then be brought to a *File Open* window. Please select the CSV file 
Your process will then run

![image](https://user-images.githubusercontent.com/98031074/155775595-48dbcafe-a4bc-4f4e-802e-b10f1a64bf04.png)

End Result: 

![image](https://user-images.githubusercontent.com/98031074/155775644-57737b62-9e68-46d1-a69e-1c0e2a03f37e.png)

Feel free to copy and past the text (Read Only)


