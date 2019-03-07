# FRC Python Scout
<h1>Current Version: 0.7
<h1>If you are testing the software please leave issue reports filed under either "bug fix" or "feature suggestions" to help with development!</h1>
<h3>This project utilizes the tbapy python library developed by FRC team 1418</h3>
<h2>Unless you are intentionally modifying or adding new features, you will only need to call event_report() function to generate a spreadsheet</h2>
<h2>Sample call to function:  event_report("2019cadm")  (Pulls information about the Del Mar weeek one regional tournament</h2>
<h2>How to Use Project in Current State:</h2>
<ul>
  <li>Ensure Microsoft Excel is installed(functionality not yet tested with LibreOffice, but may still work)</li>
  <li>Download xlwt python library (using pip install)</li>
  <li>Download tbapy python library (using pip install)</li>
  <li>Generate a new token under account page on the blue alliance</li>
  <li>Replace tba variable in code with your token as a string</li>
  <li>Call event_report(event) function with parameter of a string 'event' where event is the code for a given FRC tournament (event codes are the end of a link to an event page on the blue alliance)(event_report("2019cadm") is a working example)</li>
  <li>A new spreadsheet with basic statistics about all teams at the event will be created within the project folder</li>
</ul>
<h2>Changes to come:</h2>
<ul>
    <li>Add pre-scouting capability to search through all teams at a given event and get all history at other events for the current season</li>
</ul>

