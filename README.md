# VBA-challenge
A script that calculates and prints each tickers yearly change, yearly percentage change, and total volume for the year. It also finds the greatest percentage increase and decrease and the greatest total volume.

## About The Script
The script itself loops through all current sheets in the workbook and does its calculations on each seperate sheet. I went through a couple iterations of how it ran and with each iteration, it got more and more effencient. There are definitely more improvements I could make on the script given more time. One improvement I know I would want to make is to use checks on the dates instead of relying that the data is sorted by the ticker. 1 improvement I did make through the iterations was that in the beginning, I was running calculations and updating row values each time I was reading a new row and I improved it by running the calculations and updates only before moving to the next ticker which, obviously, cut the run time down by a huge amount.

## Screenshot

![alt text][logo]

[logo]: https://github.com/HunterG003/VBA-challenge/blob/main/images/Screenshot%202023-03-15%20at%208.29.04%20PM.png "Screenshot of Excel SHeet"
