from __future__ import print_function
from datetime import datetime, timedelta

from bs4 import BeautifulSoup
import click
from icalendar import Calendar
import pypandoc
from pytz import timezone, utc, all_timezones
import recurring_ical_events
from tqdm import tqdm
from tzlocal import get_localzone

"""Convert ICAL format into org-mode.

Files can be set as explicit file name, or `-` for stdin or stdout::

    $ python ical2org.py in.ical out.org

    $ python ical2org.py in.ical - > out.org

    $ cat in.ical | python ical2org.py - out.org

    $ cat in.ical | python ical2org.py - - > out.org

This code is a heavily-modified version of https://github.com/asoroa/ical2org.py , using
the recurring_ical_events library https://github.com/niccokunzmann/python-recurring-ical-events
to do most of the heavy lifting around handling event parsing.
"""

def org_datetime(dt, tz):
    '''
    Convert a timezone-aware datetime.datetime to YYYY-MM-DD DayofWeek HH:MM str
    in a provided timezone.
    '''
    return dt.astimezone(tz).strftime("<%Y-%m-%d %a %H:%M>")

def org_date(d, tz):
    '''Timezone aware datetime.date to YYYY-MM-DD DayofWeek in a provided timezone.
    '''
    # Convert the date to a datetime first
    dt = datetime.combine(d, datetime.min.time())
    return dt.astimezone(tz).strftime("<%Y-%m-%d %a>")

class IcalError(Exception):
    pass

class Convertor():
    def __init__(self, days=90, tz=None, include_location=True):
        """
        days: Window length in days (left & right from current time). Has
        to be positive.
        tz: timezone. If None, use local timezone.
        include_location: If False, don't add the location to
        titles of generated org entries.
        """
        self.tz = timezone(tz) if tz else get_localzone()
        self.days = days
        self.include_location = include_location

    def __call__(self, fh, fh_w):
        try:
            cal = Calendar.from_ical(fh.read())
        except ValueError as e:
            msg = "Parsing error: {}".format(e)
            raise IcalError(msg)

        now = datetime.now()
        start = now - timedelta(days=self.days)
        end = now + timedelta(days=self.days)
        events = recurring_ical_events.of(cal).between(start, end)
        for event in tqdm(events):
            summary = event["SUMMARY"]
            summary = summary.replace('\\,', ',')
            location = None
            if event.get("LOCATION", None):
                location = event['LOCATION'].replace('\\,', ',')
            if not any((summary, location)):
                summary = u"(No title)"
            else:
                summary += " - " + location if location and self.include_location else ''
            fh_w.write(u"* {}".format(summary))
            fh_w.write(u"\n")
            if isinstance(event["DTSTART"].dt, datetime):
                fh_w.write(u"  {}--{}\n".format(
                    org_datetime(event["DTSTART"].dt, self.tz),
                    org_datetime(event["DTEND"].dt, self.tz)))
            else:
                # all day event
                fh_w.write(u"  {}--{}\n".format(
                    org_date(event["DTSTART"].dt, timezone('UTC')),
                    org_date(event["DTEND"].dt - timedelta(days=1), timezone('UTC'))))
            description = event.get("DESCRIPTION", None)
            if description:
                if bool(BeautifulSoup(description, "html.parser").find()):
                    description = pypandoc.convert_text(description, "org", format="html")
                description = '\n'.join(description.split('\\n'))
                description = description.replace('\\,', ',')
                fh_w.write(u"{}\n".format(description))
            fh_w.write(u"\n")

def check_timezone(ctx, param, value):
    if (value is None) or (value in all_timezones):
        return value
    click.echo(u"Invalid timezone value {value}.".format(value=value))
    click.echo(u"Use --print-timezones to show acceptable values.")
    ctx.exit(1)

def print_timezones(ctx, param, value):
    if not value or ctx.resilient_parsing:
        return
    for tz in all_timezones:
        click.echo(tz)
    ctx.exit()


@click.command(context_settings={"help_option_names": ['-h', '--help']})
@click.option(
    "--print-timezones",
    "-p",
    is_flag=True,
    callback=print_timezones,
    is_eager=True,
    expose_value=False,
    help="Print acceptable timezone names and exit.")
@click.option(
    "--days",
    "-d",
    default=90,
    type=click.IntRange(0, clamp=True),
    help=("Window length in days (left & right from current time). "
          "Has to be positive."))
@click.option(
    "--timezone",
    "-t",
    default=None,
    callback=check_timezone,
    help="Timezone to use. (Local timezone by default).")
@click.option(
    "--location/--no-location",
    "include_location",
    default=True,
    help="Include the location (if present) in the headline. (Location is included by default).")
@click.argument("ics_file", type=click.File("r", encoding="utf-8"))
@click.argument("org_file", type=click.File("w", encoding="utf-8"))
def main(ics_file, org_file, days, timezone, include_location):
    convertor = Convertor(days, timezone, include_location)
    try:
        convertor(ics_file, org_file)
    except IcalError as e:
        click.echo(str(e), err=True)
        raise click.Abort()
