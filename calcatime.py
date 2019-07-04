#pylint: disable=line-too-long
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

"""
from enum import Enum
import re
import json
from collections import namedtuple
from datetime import datetime, timedelta
import calendar
from typing import Dict, List, Optional, Tuple


# third-party modules
from docopt import docopt


__version__ = '0.2'


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


def parse_timerange_tokens(timespan_tokens: List[str]) -> Tuple[datetime, datetime]:
    """Return start and end of the range specified by tokens."""
    # collect today info
    today = datetime.today()
    today_start = datetime(today.year, today.month, today.day, 0, 0)
    today_end = today_start + timedelta(days=1)

    # calculate this week start date
    week_start = today_start - timedelta(days=today_start.weekday())

    # count the number of times 'last' token is provided
    # remove 7 days for each count
    last_count = timespan_tokens.count('last')
    last_offset = -7 * last_count

    # count the number of times 'next' token is provided
    # add 7 days for each count
    next_count = timespan_tokens.count('next')
    next_offset = 7 * next_count

    offset = last_offset + next_offset

    # now process the known tokens
    if 'today' in timespan_tokens:
        return (today_start + timedelta(days=offset),
                today_end + timedelta(days=offset))

    elif 'yesterday' in timespan_tokens:
        return (today_start + timedelta(days=-1 + offset),
                today_end + timedelta(days=-1 + offset))

    elif 'week' in timespan_tokens:
        return (week_start + timedelta(days=offset),
                week_start + timedelta(days=7 + offset))

    elif 'month' in timespan_tokens:
        month_index = today.month + (-last_count + next_count)
        month_index = month_index if month_index >= 1 else 12
        month_start = datetime(today.year, month_index, 1)
        month_end = datetime(today.year, month_index + 1, 1) + timedelta(-1)
        return (month_start, month_end)

    elif 'year' in timespan_tokens:
        year_number = today.year + (-last_count + next_count)
        year_start = datetime(year_number, 1, 1)
        year_end = datetime(year_number + 1, 1, 1) + timedelta(-1)
        return (year_start, year_end)

    elif 'decade' in timespan_tokens:
        raise NotImplementedError()

    elif 'century' in timespan_tokens:
        raise NotImplementedError()

    elif 'millenium' in timespan_tokens:
        raise NotImplementedError()

    # process week days
    for idx, day_names in enumerate(
            zip(map(str.lower, list(calendar.day_name)),
                map(str.lower, list(calendar.day_abbr)))):
        if any(x in timespan_tokens for x in day_names):
            rstart = week_start + timedelta(days=idx + offset)
            rend = week_start + timedelta(days=idx + 1 + offset)
            return (rstart, rend)

    raise Exception('Can not determine time span.')


def get_exchange_events(server: str,
                        domain: Optional[str],
                        username: str,
                        password: str,
                        range_start: datetime,
                        range_end: datetime) -> List[CalendarEvent]:
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

    events: List[CalendarEvent] = []
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


def group_by_title(events: List[CalendarEvent]) -> Dict[str, List[CalendarEvent]]:
    """Group given events by event title."""
    grouped_events: Dict[str, List[CalendarEvent]] = {}
    for event in events:
        if event.title in grouped_events:
            grouped_events[event.title].append(event)
        else:
            grouped_events[event.title] = [event]
    return grouped_events


def group_by_category(events: List[CalendarEvent],
                      unknown_group='---') -> Dict[str, List[CalendarEvent]]:
    """Group given events by event category."""
    grouped_events: Dict[str, List[CalendarEvent]] = {}
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


def group_by_pattern(events: List[CalendarEvent],
                     pattern: str,
                     attr: str = 'title') -> Dict[str, List[CalendarEvent]]:
    """Group given events by given regex pattern and target attribute."""
    grouped_events: Dict[str, List[CalendarEvent]] = {}
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


def cal_total_duration(grouped_events: Dict[str, List[CalendarEvent]]) -> Dict[str, float]:
    """Calculate total duration of events in each group."""
    hours_per_group: Dict[str, float] = {}
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
    if not json_out | csv_out:
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
            _, pattern = grouping_attr.split(':')
            if pattern:
                grouped_events = \
                    group_by_pattern(events, pattern, attr='category')
        elif grouping_attr == 'category':
            grouped_events = \
                group_by_category(events)
        elif grouping_attr.startswith('title:'):
            _, pattern = grouping_attr.split(':')
            if pattern:
                grouped_events = \
                    group_by_pattern(events, pattern, attr='title')
        elif grouping_attr == 'title':
            grouped_events = \
                group_by_title(events)

    # prepare and dump data
    total_durations = cal_total_duration(grouped_events)
    calculated_data: List[Dict] = []
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