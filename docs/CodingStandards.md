\# GarageOS Coding Standards



Version: 1.0.0



\---



\# Objective



This document defines the official coding standards for GarageOS.



Every new function, module, documentation update, and future contribution must follow these standards.



\---



\# General Principles



\- Write clean code.

\- Write readable code.

\- Write maintainable code.

\- Avoid unnecessary complexity.

\- Prefer clarity over short code.



\---



\# Naming Conventions



\## Functions



Function names must always be descriptive.



Good examples



createCustomerRecord()



updateVehicleInformation()



calculateInventoryValue()



generateWorkOrderReport()



Bad examples



create()



save()



run()



test()



fn1()



abc()



Function names should clearly describe what they do.



\---



\## Variables



Variable names must also be descriptive.



Good



selectedCustomer



currentVehicleMileage



inventoryQuantity



paymentMethod



customerPhoneNumber



Bad



a



x



tmp



obj



data1



\---



\## Constants



Global constants must always use UPPER\_SNAKE\_CASE.



Example



CUSTOMER\_SHEET\_NAME



MAXIMUM\_DISCOUNT



DEFAULT\_LANGUAGE



API\_BASE\_URL



GOOGLE\_DRIVE\_FOLDER\_ID



\---



\## File Names



Use PascalCase.



Examples



Customers.js



Vehicles.js



Inventory.js



Payments.js



Dashboard.js



DeveloperCenter.js



\---



\# Code Organization



Each file should follow this order.



1\. Header

2\. Constants

3\. Public Functions

4\. Private Functions

5\. Helper Functions



\---



\# Function Size



Prefer functions under 80 lines.



If a function exceeds 100 lines, consider splitting it.



\---



\# Function Responsibility



Each function should perform only one responsibility.



Avoid functions that perform multiple unrelated tasks.



\---



\# Comments



Use comments only when they add value.



Avoid obvious comments.



Good



/\*\*

&#x20;\* Calculates the inventory valuation using FIFO.

&#x20;\*/



Bad



// Add one to x



x++;



\---



\# Documentation



Every public function should include documentation.



Example



/\*\*

&#x20;\* Creates a new customer.

&#x20;\*

&#x20;\* @param {Object} customerData

&#x20;\* @returns {Object}

&#x20;\*/



\---



\# Error Handling



Always validate inputs.



Throw meaningful errors.



Never silently ignore exceptions.



\---



\# Logging



Use Logger only for debugging.



Remove unnecessary logs before production.



\---



\# Formatting



Indentation



4 spaces.



Maximum line length



120 characters.



Leave one blank line between functions.



\---



\# Git Commits



Commit messages should be descriptive.



Good



Add customer validation module



Fix inventory calculation bug



Implement payment history



Bad



Update



Fix



Test



Changes



\---



\# Branches



main



Production-ready code.



develop



Integration branch.



feature/module-name



New features.



bugfix/module-name



Bug fixes.



\---



\# Folder Structure



appscript/



Google Apps Script source code.



docs/



Technical documentation.



assets/



Project assets.



screenshots/



Images.



videos/



Demo videos.



tests/



Testing resources.



\---



\# Code Style



Prioritize readability over fewer lines of code.



Avoid abbreviations.



Prefer long descriptive names.



\---



\# Emojis



Do not use emojis in:



Code



Comments



Documentation



Commit messages



Technical documents



\---



\# GarageOS Philosophy



The code should be understandable six months later without additional explanations.



If a function name needs explanation, rename the function.



Code is written once but read many times.

