/** A template for a calendar event with start and end times. */
type Event = {
    title: string;
    startTime: Date;
    endTime: Date;
}

/** A template for an all-day calendar event. */
type AllDayEvent = {
    title: string;
    date: Date;
}

/** A convenient alias for all the types returned by a Google Form ItemResponse. */
type Response = string | string[] | string[][]

/** Dummy function for authorizing scopes. */
function authorize() {
    FormApp.getActiveForm(); // Google doesn't pick up the need for this scope.
    Logger.log("Everything is fine and dandy!");
}

/** The main entry point. Is attached to a trigger for the form. */
function onFormSubmit(event: GoogleAppsScript.Events.FormsOnFormSubmit) {
    let calendar = getExtraBoardCalendar();
    if (calendar === undefined) {
        Logger.log("Calendar does not exist. Exiting.")
        return;
    }

    let newEvents = Array.from(makeEvents(event.response.getItemResponses()));
    if (newEvents.length > 0) {
        // Clear any events on the selected day.
        let first = newEvents[0];
        if ("startTime" in first)
            deleteEventsThatStartOnDay(calendar, first.startTime);
        else if ("date" in first)
            deleteEventsThatStartOnDay(calendar, first.date);

        // Create the new events.
        for (let evt of newEvents) {
            if ("startTime" in evt)
                calendar.createEvent(evt.title, evt.startTime, evt.endTime);
            else if ("date" in evt)
                calendar.createAllDayEvent(evt.title, evt.date);
        }
    }
}

/** Delete any events that start on the given day. */
function deleteEventsThatStartOnDay(calendar: GoogleAppsScript.Calendar.Calendar, date: Date) {
    let events = calendar.getEventsForDay(date);
    for (let evt of events.filter((e) => {
        let start = new Date();
        start.setTime(e
            .getStartTime()
            .getTime());
        return dateIsWithinDay(start, date);
    })) {
        evt.deleteEvent();
    }
}

/** Determine whether a particular time falls within the range of a day. */
function dateIsWithinDay(date: Date, day: Date): boolean {
    let todayStart = new Date(date);
    setTimeComponent(todayStart, { hours: 0, minutes: 0 });
    let tomorrowStart = addToDate(date, 24 * 60 * 60 * 1000);
    setTimeComponent(tomorrowStart, { hours: 0, minutes: 0 });
    let dateTime = date.getTime();
    return dateTime >= todayStart.getTime() && dateTime < tomorrowStart.getTime();
}

/** Gets the user's Extra Board calendar. */
function getExtraBoardCalendar(): GoogleAppsScript.Calendar.Calendar {
    let calendars = CalendarApp.getOwnedCalendarsByName("GET Bus");
    if (calendars.length > 0)
        return calendars[0];
}

/** Build calendar event(s) for the day's assignment. */
function* makeEvents(ir: GoogleAppsScript.Forms.ItemResponse[]): Iterable<Event | AllDayEvent> {
    let responses = ir.map((r) => r.getResponse());
    let generator: (r: (Response)[]) => Iterable<Event | AllDayEvent>
    switch (responses[0]) {
        case "Have a run":
            generator = makeRunEvents;
            break;
        case "On show":
            generator = makeShowEvents;
            break;
        case "Day off!":
            generator = makeDayOffEvents;
            break;
    }
    yield* Array.from(generator(responses));
}

function* makeRunEvents(r: Response[]): Iterable<Event | AllDayEvent> {
    // Load runs from the Post Bid Report.
    let pbr = SpreadsheetApp
        .openByUrl("https://docs.google.com/spreadsheets/d/1hqlE6HPQFAFQ-DuaXvOZDgXTgfTqaUCHbTlAgOjF1-k/edit")
        .getSheetByName("Post Bid Report")
        .getDataRange()
        .offset(1, 0);
    let runs = loadBiddedRuns(pbr);

    let runNumber = r[1];
    if (typeof runNumber === "string" && runs.has(runNumber)) {
        let run = runs.get(runNumber);
        let date = makeDate(r[2]);
        if ("shift" in run) { // straight run
            let range = makeDateRange(date, run.shift.reportTime, run.shift.signOut);
            yield {
                title: `Run ${runNumber} Block ${run.shift.block}`,
                startTime: range[0],
                endTime: range[1]
            }
        } else { // split run
            let firstRange = makeDateRange(date, run.firstHalf.reportTime, run.firstHalf.signOut);
            yield {
                title: `Run ${runNumber} Block ${run.firstHalf.block}`,
                startTime: firstRange[0],
                endTime: firstRange[1]
            }

            let secondRange = makeDateRange(date, run.secondHalf.reportTime, run.secondHalf.signOut);
            yield {
                title: `Run ${runNumber} Block ${run.secondHalf.block}`,
                startTime: secondRange[0],
                endTime: secondRange[1]
            }
        }
    } else {
        Logger.log(`Run does not exist: ${runNumber}`)
    }
}

function* makeShowEvents(r: Response[]): Iterable<Event | AllDayEvent> {
    let time = parseResponseTime(r[1]);
    let secondTime = parseResponseTime(r[2]);
    let date = makeDate(r[3]);
    if (secondTime === undefined) { // straight show
        let start = new Date(date);
        setTimeComponent(start, time);
        let end = addToDate(start, 8 * 60 * 60 * 1000);
        yield { title: "On Show", startTime: start, endTime: end };
    } else { // split show
        let firstStart = new Date(date);
        setTimeComponent(firstStart, time);
        let firstEnd = addToDate(firstStart, 4 * 60 * 60 * 1000);
        yield { title: "On Show", startTime: firstStart, endTime: firstEnd };

        let secondStart = new Date(date);
        setTimeComponent(secondStart, secondTime);
        let secondEnd = addToDate(secondStart, 4 * 60 * 60 * 1000);
        yield { title: "On Show", startTime: secondStart, endTime: secondEnd };
    }
}

function* makeDayOffEvents(r: Response[]): Iterable<Event | AllDayEvent> {
    let date = makeDate(r[1]);
    yield { title: "Day Off", date: date };
}

/** Parse a time entered into a Google Forms time control. */
function parseResponseTime(response: Response): Time {
    if (typeof response === "string") {
        let match = response.match(/(\d\d?):(\d\d?)/);
        if (match == null)
            return undefined;
        else
            return { hours: parseInt(match[1]), minutes: parseInt(match[2]) };
    } else {
        return undefined;
    }
}

/** Instantiate a time range, which is allowed to cross the midnight boundary. */
function makeDateRange(date: Date, from: Time, to: Time): Date[] {
    let wrapsAround = to.hours * 60 + to.minutes < from.hours * 60 + from.minutes;
    let start = new Date(date);
    setTimeComponent(start, from);
    let end = wrapsAround ? addToDate(date, 24 * 60 * 60 * 1000) : new Date(date);
    setTimeComponent(end, to);
    return [start, end];
}

/** Read the user's choice of today or tomorrow. */
function makeDate(response: Response): Date {
    let today = new Date();
    switch (response) {
        case "Tomorrow":
            return addToDate(today, 24 * 60 * 60 * 1000);
        case "Today":
        default:
            return today;
    }
}

/** Set the hour and minute components of a Date. */
function setTimeComponent(date: Date, time: Time) {
    date.setHours(time.hours);
    date.setMinutes(time.minutes);
    date.setSeconds(0);
    date.setMilliseconds(0);
}

/** Copy a Date, then add the provided offset to the copy. */
function addToDate(date: Date, milliseconds: number): Date {
    let d = new Date();
    d.setTime(date.getTime() + milliseconds);
    return d;
}