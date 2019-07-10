# calcatime
python utility script that collects total time from categorized calendar events


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

## Examples

```bash
$ calcatime -c "office365" -u "myemail@mycomp.com" -p $password last week --json
```