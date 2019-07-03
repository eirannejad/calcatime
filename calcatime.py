"""Calculates totat time from calendar events and groupes by an event attribute.

Usage:
    calcatime -c <calendar_uri> [-d <domain>] -u <username> -p <password> <timespan>... [--by <event_attr>] [--include-zero] [--json|--csv]

Options:
    -h, --help              Show this help
    -V, --version           Show command version
    -c <calendar_uri>       Calendar provider:server uri
                            See Calendar Providers
    -d <domain>             Domain name
    -u <username>           User name
    -p <password>           Password
    <timespan>              Only include events in given time span
                            See Timespan Options
    --by=<event_attr>       Group total times by given event attribute
                            See Event Attributes
    --include-zero          Include zero totals in output
    --json                  Output data to json
    --csv                   Output data to csv

Calendar Providers:
    Microsoft Exchange:     exchange:<server url>

Timespan Options:
    today
    yesterday
    monday | mon
    tuesday | tue
    wednesday | wed
    thurday | thu
    friday | fri
    saturday | sat
    sunday | sun
    week (current week)
    last (can be used multiple times e.g. last last week)

Event Grouping Attributes:
    category[:<regex_pattern>]
    title[:<regex_pattern>]

"""
from enum import Enum
import re
import json
from collections import namedtuple
from datetime import datetime, timedelta

# third-party modules
from docopt import docopt


__version__ = '0.1'


DATETIME_FORMAT = '%Y-%m-%d'

# enum for calendar providers
class CalendarProvider(Enum):
    """Supported calendar providers"""
    Exchange = 0


# unified tuple for holding calendar event properties
CalendarEvent = namedtuple('CalendarEvent', [
    'title',
    'start',
    'end',
    'duration',
    'categories'
])


def parse_timerange_tokens(timespan_tokens):
    """Return start and end of the range specified by tokens."""
    # collect today info
    today = datetime.today()
    today_start = datetime(today.year, today.month, today.day, 0, 0)
    today_end = today_start + timedelta(days=1)

    # calculate this week start date
    week_start = today_start - timedelta(days=today_start.weekday())

    # count the number of times 'last' token is provided
    # remove 7 days for each time
    last_offset = -7 * timespan_tokens.count('last')

    # now process the tokens
    if 'today' in timespan_tokens:
        return (today_start + timedelta(days=last_offset),
                today_end + timedelta(days=last_offset))
    elif 'yesterday' in timespan_tokens:
        return (today_start + timedelta(days=-1 + last_offset),
                today_end + timedelta(days=-1 + last_offset))

    if 'monday' in timespan_tokens or 'mon' in timespan_tokens:
        offset_range = (0 + last_offset, 1 + last_offset)
    elif 'tuesday' in timespan_tokens or 'tue' in timespan_tokens:
        offset_range = (1 + last_offset, 2 + last_offset)
    elif 'wednesday' in timespan_tokens or 'wed' in timespan_tokens:
        offset_range = (2 + last_offset, 3 + last_offset)
    elif 'thursday' in timespan_tokens or 'thu' in timespan_tokens:
        offset_range = (3 + last_offset, 4 + last_offset)
    elif 'friday' in timespan_tokens or 'fri' in timespan_tokens:
        offset_range = (4 + last_offset, 5 + last_offset)
    elif 'saturday' in timespan_tokens or 'sat' in timespan_tokens:
        offset_range = (5 + last_offset, 6 + last_offset)
    elif 'sunday' in timespan_tokens or 'sun' in timespan_tokens:
        offset_range = (6 + last_offset, 7 + last_offset)
    elif 'week' in timespan_tokens or 'last' in timespan_tokens:
        offset_range = (0 + last_offset, 7 + last_offset)

    if offset_range[0] >= 0:
        rstart = week_start + timedelta(days=offset_range[0])
    else:
        rstart = week_start - timedelta(days=abs(offset_range[0]))

    if offset_range[1] >= 0:
        rend = week_start + timedelta(days=offset_range[1])
    else:
        rend = week_start - timedelta(days=abs(offset_range[1]))

    return (rstart, rend)


def get_exchange_events(server, domain, username, password,
                        range_start, range_end):
    """Connect to exchange calendar server and get events within range."""
    # load exchange module if necessary
    from exchangelib import Credentials, Configuration, Account, DELEGATE
    from exchangelib import EWSDateTime, EWSTimeZone

    # setup access
    account = Account(
        primary_smtp_address=username,
        config=Configuration(
            server=server,
            credentials=Credentials(
                r'{}\{}'.format(domain, username),
                password
                )
            ),
        autodiscover=False,
        access_type=DELEGATE
        )

    events = []
    localzone = EWSTimeZone.localzone()
    local_start = localzone.localize(EWSDateTime.from_datetime(range_start))
    local_end = localzone.localize(EWSDateTime.from_datetime(range_end))
    for item in account.calendar.filter(    ##pylint: disable=no-member
            start__range=(local_start, local_end)).order_by('start'):
        events.append(
            CalendarEvent(
                title=item.subject,
                start=item.start,
                end=item.end,
                duration=(item.end - item.start).seconds / 3600,
                categories=item.categories
            ))
    return events


def group_by_title(events):
    grouped_events = {}
    for event in events:
        if event.title in grouped_events:
            grouped_events[event.title].append(event)
        else:
            grouped_events[event.title] = [event]
    return grouped_events


def group_by_category(events, unknown_group='---'):
    grouped_events = {}
    for event in events:
        if event.categories:
            for cat in event.categories:
                if cat in grouped_events:
                    grouped_events[cat].append(event)
                else:
                    grouped_events[cat] = [event]
        else:
            if unknown_group in grouped_events:
                grouped_events[unknown_group].append(event)
            else:
                grouped_events[unknown_group] = [event]
    return grouped_events


def group_by_pattern(events, pattern, attr='title'):
    grouped_events = {}
    for event in events:
        target_tokens = []
        if attr == 'title':
            target_tokens.append(event.title)
        elif attr == 'category':
            target_tokens = event.categories

        if target_tokens:
            for token in target_tokens or []:
                match = re.search(pattern, token, flags=re.IGNORECASE)
                if match:
                    matched_token = match.group()
                    if matched_token in grouped_events:
                        grouped_events[matched_token].append(event)
                    else:
                        grouped_events[matched_token] = [event]
                    break

    return grouped_events


def cal_total_duration(grouped_events):
    hours_per_group = {}
    for event_group, events in grouped_events.items():
        total_duration = 0
        for event in events:
            total_duration += event.duration
        hours_per_group[event_group] = total_duration
    return hours_per_group


def main():
    """Parse arguments, parse time span, get and organize events, dump data."""
    # process command line args
    args = docopt(__doc__, version='calcatime {}'.format(__version__))

    # determine calendar provider
    calprovider = calserver = None
    calendar_uri = args.get('-c', None)
    if calendar_uri:
        if calendar_uri.startswith('exchange:'):
            calserver = calendar_uri.replace('exchange:', '')
            calprovider = CalendarProvider.Exchange
    else:
        raise Exception('Calendar provider is required.')

    # determine credentials
    username = args.get('-u', None)
    password = args.get('-p', None)
    if not username or not password:
        raise Exception('Calendar access credentials are required.')

    # get domain if provided
    domain = args.get('-d', None)

    # determine grouping attribute, set defaults if not provided
    grouping_attr = args.get('--by', None)
    if not grouping_attr:
        if calprovider == CalendarProvider.Exchange:
            grouping_attr = 'category'
        else:
            grouping_attr = 'title'

    # determine if zeros need to be included
    include_zero = args.get('--include-zero', False)

    # determine output type, set defaults if not provided
    json_out = args.get('--json', False)
    csv_out = args.get('--csv', False)
    if not (json_out | csv_out):
        csv_out = True

    # determine requested time span
    start, end = parse_timerange_tokens(
        args.get('<timespan>', [])
    )

    # collect events from calendar
    events = None
    if calprovider == CalendarProvider.Exchange:
        events = get_exchange_events(
            server=calserver,
            domain=domain,
            username=username,
            password=password,
            range_start=start,
            range_end=end
            )
    else:
        raise Exception('Unknown calendar provider.')

    # group events
    grouped_events = {}
    if events:
        if grouping_attr.startswith('category:'):
            token, pattern = grouping_attr.split(':')
            if pattern:
                grouped_events = \
                    group_by_pattern(events, pattern, attr='category')
        elif grouping_attr == 'category':
            grouped_events = \
                group_by_category(events)
        elif grouping_attr.startswith('title:'):
            token, pattern = grouping_attr.split(':')
            if pattern:
                grouped_events = \
                    group_by_pattern(events, pattern, attr='title')
        elif grouping_attr == 'title':
            grouped_events = \
                group_by_title(events)

    # prepare and dump data
    total_durations = cal_total_duration(grouped_events)
    calculated_data = []
    for event_group, events in grouped_events.items():
        if not include_zero and total_durations[event_group] == 0:
            continue
        calculated_data.append({
            'start': start.strftime(DATETIME_FORMAT),
            'end': end.strftime(DATETIME_FORMAT),
            'group': event_group,
            'duration': total_durations[event_group]
        })

    if json_out:
        print(json.dumps(calculated_data))
    elif csv_out:
        print('"start","end","group","duration"')
        for data in calculated_data:
            print(','.join([
                '"{}"'.format(data['start']),
                '"{}"'.format(data['end']),
                '"{}"'.format(data['group']),
                str(data['duration'])
            ]))


if __name__ == '__main__':
    main()
