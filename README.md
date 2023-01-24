# OutlookHelper
## Description
This is a side project meant to ease the process of filling in daily work ours in your outlook calendar (an easy way to centralize emails and work done on a daily basis) and forward them to any company project management tool.

This current helper lets you retrieve your outlook calendar events and push relevant hours to an online project management tool (e.g. Redmine).

## Prerequisite
You must have Outlook installed on your PC in order to use the tool (tested on version 2212).

## App configuration
The configuration of the application is done through the appsettings.json file. Here is an explanation of each field of the file:
```json
{
  "ExploratorConfiguration": { // the config for the outlook explorator
    "WorkingPercentage": "1", // your work percentage (from 0% to 100%)
    "ExcludedCategories": [ "<excludedCategory>" ], // outlook category to exclude while exploring
    "ExcludedSubjects": [ "<excluded subject 1>", "<excluded subject 2>" ], // outlook events subjects to exclude while exploring
    "WeekRangePerYear": [ // allow to filter the exploration of your outlook calendar per week range and year
      {
        "Year": "2022",
        "WeekRange": [ "1", "52" ]
      },
      {
        "Year": "2023",
        "WeekRange": [ "1", "52" ]
      }
    ]
  },
  "ExporterConfiguration": { // the configuration for the redmine exporter
    "ServerUrl": "<serverUrl>", // the redmine server url
    "ApiKey": "<apiKey>", // your own api-key, available in your account information
    "TargetUserId": "0", // your user id on redmine (integer)
    "TimeEntryActivities": [ // the activities in redmine which can be matched to the outlook calendar categories
      {
        "Name": "<activityName>",
        "Id": 0
      }
    ],
    "CustomCategoryToActivityMappings": [ // custom mapping from outlook category to a redmine activity
      {
        "Category": "<category>",
        "ActivityId": 0
      }
    ]
  }
}
```
## How to use the tool
The OutlookHelper tool assumes that you enter all the information of a desired time entry through an outlook event (or appointment), namely:
- A description of the task that was done in the subject of the outlook event, including the corresponding Redmine issue indicated by a hashtag followed by the issue number (e.g. <i>'Meeting with the A team - #0000'</i>),
- A categories linked to that outlook event, matching by name an activity on Redmine (hint: you can also associate colors with categories in outlook, so that you have a better readibility of your weekly schedule)

The OutlookHelper tool will ignore all tasks that are not associated to any category while exploring, thus they will also be unavailable for export.

In the particular case of meetings, you might want to link a Redmine issue to it, while not resending the meeting subject update to all participants. Currently, you will have to use the following workaround:
- Assign a category to the meeting that you can then exclude from the exploration using the appsettings.json file described above,
- Add a new appointment describing the meeting at the same time and duration, and assign it an explorable category. The tool will then be able to export only this particular event describing the meeting.