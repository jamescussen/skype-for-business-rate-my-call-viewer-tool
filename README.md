Skype for Business Rate My Call Viewer Tool
===========================================

            

The Skype for Business Rate My Call Viewer Tool is a Powershell based GUI interface for listing up and graphing Rate My Call data out of the Skype for Business Monitoring database. Currently there is no Monitoring report for accessing this data so I created
 this tool to help fill the gap. 


![Image](https://github.com/jamescussen/skype-for-business-rate-my-call-viewer-tool/raw/master/ratemycallviewer1.01_sm.png)


**Features:**


  *  Select your required start and end date, rating filter, SIP URI filter, Reason filter, and Voice and/or Video to be listed up. Note the filters use regex, so for example you could use it to filter for multiple reasons using the OR Operator like this “echo|backgroundnoise”.

  *  Export all events into CSV format. 
  *  Create graphs (Stars Bar Graph, Stars Pie Graph, Reason Pie Graph, Reason Bar Graph, Type Pie Graph, Stars Stacked Bar Graph, Trend Over Time Line) of your Call Rating data.


**Version 1.01 Update:**


  *  Added the ability to select individual monitoring servers from a drop down box. This was added for large environments that have multiple monitoring databases and only want to retrieve statistics from one at a time. By default, all monitoring database will
 be queried. 
  *  Added a check for the database version. The tool only works on Skype for Business databases so the check makes sure the database is at least Version 7 (ie. Skype for Business level).


**Version 1.02 Update (20/04/2017):**


  *  Added date/time localisation checkbox. By default the monitoring server records time is in GMT. This update adds a checkbox to localise all the date/time values to be in the timezone of the server you are running it on (instead of GMT). This changes the
 date pickers as well as the date displayed in the list and graphs. 
  *  Added the ability to zoom in on the Trend Over Time chart. You do this by clicking and dragging the mouse on the area of the graph you want to zoom to, scroll bars will appear so you can scroll the zoomed in view.


**Version 1.03 Update (22/04/2017)**


  *  Fixed an issue with the SQL query used for Video / Audio. The query now gets all records.

  *  Fixed issue with data grid view scroll bar refresh. 
  *  Fixed a sorting issue with the Stacked Bar and Trend Over Time Graphs that would cause an issue with the output.

  *  More accurate graphs! When both video and voice are selected the rating data gets listed twice for each call because video calls contain both voice and video ratings. So in previous versions the star ratings were counted as separate calls which artificially
 inflated the star rating value given. In this version the double counting of this data has been removed from star rating graphs, with the voice and video star rating given by each user being combined. 


**Version 1.04 (15/5/2017) – C2R Update**


  *  Now Supports Skype for Business C2R 2016 client Rate My Call issue items. The C2R 2016 client has an entirely new set of rate my call feedback, so the tool has been updated to include these.

  *  Re-worked the graphs again to handle new data 
  *  Voice and Video calls don't get listed twice in this version (as it did in the previous version), graph processing was updated from previous version to handle this.

  *  Get Records processing speed was increased by limiting records by date range in SQL query.

 


**1.05 Update (16/3/2018)**

  *  Total Rows Counter added at the bottom 
  *  'Top 10 One Star Users' graph added. This can be used so you can follow up with these users about their bad experiences.

  *  'Top 10 Zero Star Users (Lync 2013 Client)' graph added. This can be used to follow up on Lync 2013 client users that are not responding the Rate My Call dialog.


 


**Prerequisites:**


  *  This tool should be run on a machine that has the Skype for Business powershell module installed. This is required because the 'Get-CSService' command is used to discover the location of the Montoring Database.

  *  The user running the tool needs to have sufficient rights to run select queries on the 'QoEMetrics' database and SELECT access on the following tables: Session, AudioStream, CallQualityFeedback, CallQualityFeedbackToken, CallQualityFeedbackTokenDef, User,
 MediaLine 

 


**Example Graph:**


![Image](https://github.com/jamescussen/skype-for-business-rate-my-call-viewer-tool/raw/master/Graph-StarBar_sm.png)


 


**For full information on this tool see the following link: [http://www.myteamslab.com/2017/01/skype-for-business-rate-my-call-viewer.html](http://www.myteamslab.com/2017/01/skype-for-business-rate-my-call-viewer.html)**


 





        
    
