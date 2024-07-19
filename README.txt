#Defined the Dim variables

#Created the headers for the new columns 

#Created two separate LastRow's due to different calculations 

#Initial LastRow is used for the entire data set to run the initial calculation

#Calculation Steps:
	1) Created the list of all Tickers and placed in column I
	2) Ran Quarterly Change Calculation and placed in column J
	3) Ran Percent Change Calculation and placed in column K
	4) Ran the Total Vol Calculation and placed in column L
	5) Formatted the Quarterly Change based on their values if less 0 red, 
           greater than green and if zero then blank

#After the Calculation, the Greatest % Increase and Decrease and Max Vol values are found
	1) Defined another LastRow since only considering the new list that was created
	2) Looped through the percent changes column to find the greatest increase/decrease 
	   and Max Volume
	
Since there are 4 sheets for Q1-Q4, the loop added to apply the same calculation across other sheets
	

	