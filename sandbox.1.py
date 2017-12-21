import datetime
import dateutil.rrule as rrule

def next_business_day:

    holidays = [
        datetime.datetime(2017, 12, 21,),
        datetime.datetime(2012, 6, 1,),
        # ...
    ]

    # Create a rule to recur every weekday starting today
    r = rrule.rrule(rrule.DAILY,
                    byweekday=[rrule.MO, rrule.TU, rrule.WE, rrule.TH, rrule.FR],
                    dtstart=datetime.date.today())

    # Create a rruleset
    rs = rrule.rruleset()

    # Attach our rrule to it
    rs.rrule(r)

    # Add holidays as exclusion days
    for exdate in holidays:
        rs.exdate(exdate)


    print rs[0]