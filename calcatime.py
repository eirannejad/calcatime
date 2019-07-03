"""Calculates totat time from calendar events.

Usage:
    calcatime (-h | --help)
    calcatime (-V | --version)
    calcatime -c <calendar_uri> [-d <domain>] -u <username> -p <password> <timespan>... [--json|--csv]

Options:
    -h, --help              Show this help
    -V, --version           Show command version
    -c <calendar_uri>       Calendar provider:server uri. See Providers.
    -d <domain>             Domain name
    -u <username>           User name
    -p <password>           Password
    --json                  Output data to json
    --csv                   Output data to csv

Calendar Providers:
    Microsoft Exchange:     exchange:<server url>

Timespan Tokens:
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
    last (can specify multiple times e.g. last week)
"""
#pylint: disable=C0103
from enum import Enum
from collections import namedtuple
from datetime import datetime, timedelta

from docopt import docopt


__version__ = '1.0'


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


def get_timecal_range(timespan_tokens):
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
    tz = EWSTimeZone.localzone()
    local_start = tz.localize(EWSDateTime.from_datetime(range_start))
    local_end = tz.localize(EWSDateTime.from_datetime(range_end))
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


def main():
    """Parse arguments, parse time span, get and organize events, dump data."""
    # process command line args
    args = docopt(__doc__, version='calcatime {}'.format(__version__))

    # determine calendar provider
    calserver = None
    calprovider = args.get('-c', None)
    if calprovider:
        if calprovider.startswith('exchange:'):
            calserver = calprovider.replace('exchange:', '')
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

    # determine requested time span
    start, end = get_timecal_range(
        args.get('<timespan>', [])
    )

    # collect events from calendar
    if calprovider == CalendarProvider.Exchange:
        events = get_exchange_events(
            server=calserver,
            domain=domain,
            username=username,
            password=password,
            range_start=start,
            range_end=end
            )

    # organize events
    for event in events:
        print(event)

    # dump data


if __name__ == '__main__':
    main()
