hi, and thank you for downloading the mailhaus.xlsx mailing machine to meet all your 
reverse geocoding needs for optimal direct mailing campaigns and the overall reduction of mail merge errors.

	
you will find a number of files in this repository, so allow me to walk you 
through what exactly it is you're looking at.


the first is a mailing master sheet excel file. this will allow you to begin manually 
entering contact, property, and mailing information. 


the primary purpose of this repository is to streamline the infill of missing geographical information 
for properties and mailing addresses of interest specific to your clients.


once you have entered as much information as you can, reference the sheet titled "Overview" to see how
many records/rows have been deemed as incomplete.


select ALT + F11 on your keyboard to begin creating the mailing machine.


once the vba interface is opened, select insert and click "module"


go into the mailhaus.xlsx repository and copy the code from each individual file into a new module.


	* when creating modules from the txt files in the GeoCode API folder, make sure to create a free Google Cloud
	account and enable access to a Google Maps GeoCoding API Key that you will insert into the function in place of
	the current text that reads "YOUR API KEY HERE"


once all modules have been loaded in, click on the Developer tab of your Excel ribbon, and select "Macros"


simply select the tasks that you would like to perform, click "Run", and they will infill/format the corresponding data to your needs.
