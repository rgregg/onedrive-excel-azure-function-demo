# OneDrive, Excel REST Azure Function Demo

This repo contains the source for an Azure Function that leverages Microsoft Graph, OneDrive, and Excel REST API to provide real time data updates to an Excel workbook saved in OneDrive for Business.

## How it works

The sample uses Webhooks to receive changes from OneDrive for Business, Microsoft Graph and OneDrive API to retrieve a set of files that have changed, and the Excel Workbook API to find requests for data inside a workbook.
The sample then parses those requests, uses other APIs to retrieve data relevent to the request, and populates the workbook with this data.

End to end, the scenario allows you to use the phrase `!roland` to request stock quotes for any stock ticker symbol without leaving Excel Web App.

## Learn more

**TODO Documentation**
