---
page_type: sample
products:
- office-sp
languages:
- javascript
- typescript
extensions:
  contentType: samples
  technologies:
  - SharePoint Framework
  createdDate: 5/22/2020 12:00:00 AM
---

# Modern Calendar

## spfx-modern-calendar

This Webpart plot the SharePoint Calendar Events from the Sharepoint list (i.e., Custom List or Calendar List) to SPFx webpart based on the selected WebPart properties.

### Setup the solution and WebPart

-   Clone the Solution repo
-   Run the below commands
    -   npm install
    -   gulp build
    -   gulp serve
-   Open the SharePoint Online site workbatch (i.e., <SharePoint Site URL>/_layouts/workbench.aspx)
-   Setup below webpart properties
    -   WebPart Information
        -   Webpart Title
        -   Event Background Color
        -   Event Title Color
    -   List Information
        -   List Title (i.e., This dropdown will populate the calendar and custom sharepoint list of current site)
        -   Start Date Field (i.e., This will populate the available Date time fields of current selected List)
        -   End Date Field (i.e., This will populate the available Date time fields of current selected List)
        -   Event Title Field (i.e., This will populate the available Single Line Text fields of current selected List)
        -   Event Description Field (i.e., This will populate the available Multiple Line Text fields of current selected List)
        -   All Day Event Field (i.e., This will populate the available Boolean and AllDayEvent type fields of current selected List)
        -   Display Form URL (i.e., Provide the Selected list Display Form URL, this field is usefull to setup the View Event button in the Event popup)
        -   Show Recurrence Events (i.e., This checkbox is usefull to hide/show the recurrence events)
            -   Note: This field only available for the calendar type of list

![SS1](https://github.com/SohoDragon/SohoCalendar/blob/master/Documents/spfx-moderncalendar.gif)

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources

### Build Package

-   gulp clean
-   gulp build
-   gulp bundle --ship
-   gulp package-solution --ship

### Features

-   Configurable List (i.e., Can use Custom or Calendar Type List)
-   Configurable Event Background and Title color
-   Hide/Show Recurrence Events

### Supports

-   IE11+, Chrome, Microsoft Edge, Mozilla Firefox, Mobile

### Solution

Solution|Author(s)
--------|---------
spfx-moderncalendar | Navneet Bhimani (SOHO Dragon)

### Version history

Version|Date|Comments
-------|----|--------
1.0.0.0|May 22, 2020|Initial release
