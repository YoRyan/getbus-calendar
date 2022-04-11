/** Simple container for a time at any part of the day. */
type Time = {
    hours: number;
    minutes: number;
}

/** Represents a single work run. */
type Piece = {
    run: string;
    block: string;
    reportTime: Time;
    signOut: Time;
}

/** Represents a straight shift. */
type StraightRun = {
    shift: Piece;
}

/** Represents a split shift. */
type SplitRun = {
    firstHalf: Piece;
    secondHalf: Piece;
}

/** Parse a time entered into the Post Bid Report spreadsheet. */
function parseBidReportTime(value: string): Time | undefined {
    let match = value.match(/(\d\d?):(\d\d?)/);
    if (match == null)
        return undefined;
    else
        return { hours: parseInt(match[1]), minutes: parseInt(match[2]) };
}

function* _readBiddedShifts(pbr: GoogleAppsScript.Spreadsheet.Range): Iterable<Piece> {
    for (let row of pbr.getDisplayValues()) {
        let reportTime = parseBidReportTime(row[2]);
        let signOut = parseBidReportTime(row[3]);
        if (reportTime !== undefined && signOut !== undefined)
            yield { run: row[0], block: row[1], reportTime: reportTime, signOut: signOut }
    }
}

/** Load all runs from the Post Bid Report. */
function loadBiddedRuns(pbr: GoogleAppsScript.Spreadsheet.Range): Map<string, StraightRun | SplitRun> {
    // Group shifts by common run number.
    let groupedShifts = new Map<string, Piece[]>();
    for (let shift of Array.from(_readBiddedShifts(pbr))) {
        if (groupedShifts.has(shift.run)) {
            let arr = groupedShifts.get(shift.run);
            arr.push(shift);
        } else {
            groupedShifts.set(shift.run, [shift]);
        }
    }

    // Categorize the runs into straights and splits.
    let runs = new Map<string, StraightRun | SplitRun>();
    for (let [run, shifts] of Array.from(groupedShifts)) {
        switch (shifts.length) {
            case 1:
                runs.set(run, { shift: shifts[0] });
                break;
            case 2:
                runs.set(run, { firstHalf: shifts[0], secondHalf: shifts[1] });
                break;
        }
    }
    return runs;
}