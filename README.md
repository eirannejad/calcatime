# calcatime

This is a python commandl ine utility that helps collecting calendar events from a calendar provider (e.g. Office365) within the specified period, grouping the events using an event attribute, and calculating the total time spent for each group.

The calendar shown below, assigned categories to each event and this utility can be used to collect the events, group by category, and output the total time spent under each category.

![](art/calendar.png)

## Examples

Outputing to CSV

```bash
$ calcatime -c "office365" -u "myemail@mycomp.com" -p $password last week
```

```
"start","end","group","duration"
"2019-07-01","2019-07-08","Docs",3.0
"2019-07-01","2019-07-08","BizDev",1.0
"2019-07-01","2019-07-08","Training",13.25
"2019-07-01","2019-07-08","Standards",4.0
"2019-07-01","2019-07-08","Data",8.5
```

Outputing to JSON

```bash
$ calcatime -c "office365" -u "myemail@mycomp.com" -p $password last week --json
```

```json
[{
    "start": "2019-07-01",
    "end": "2019-07-08",
    "group": "Docs",
    "duration": 3.0
}, {
    "start": "2019-07-01",
    "end": "2019-07-08",
    "group": "BizDev",
    "duration": 1.0
}, {
    "start": "2019-07-01",
    "end": "2019-07-08",
    "group": "Training",
    "duration": 13.25
}, {
    "start": "2019-07-01",
    "end": "2019-07-08",
    "group": "Standards",
    "duration": 4.0
}, {
    "start": "2019-07-01",
    "end": "2019-07-08",
    "group": "Data",
    "duration": 8.5
}]
```

## Install

```
pip install calcatime
```

## Usage

```
$ calcatime -c <calendar_uri> [-d <domain>] -u <username> -p <password> <timespan>... [--by <event_attr>] [--include-zero] [--json] [--debug]

Options:
    -h, --help              Show this help
    -V, --version           Show command version
    -c <calendar_uri>       Calendar provider:server uri
                            See Calendar Providers ↓
    -d <domain>             Domain name
    -u <username>           User name
    -p <password>           Password
    <timespan>              Only include events in given time span
                            See Timespan Options ↓
    --by=<event_attr>       Group total times by given event attribute
                            See Event Attributes
    --include-zero          Include zero totals in output
    --json                  Output data to json; default is csv
    --debug                 Extended debug logging

Calendar Providers:
    Microsoft Exchange:     exchange:<server url>
    Office365:              office365[:<server url>]
                            default server url = outlook.office365.com

Timespan Options:
    today
    yesterday
    week (current)
    month (current)
    year (current)
    monday | mon
    tuesday | tue
    wednesday | wed
    thurday | thu
    friday | fri
    saturday | sat
    sunday | sun
    last (can be used multiple times e.g. last last week)

Event Grouping Attributes:
    category[:<regex_pattern>]
    title[:<regex_pattern>]
```