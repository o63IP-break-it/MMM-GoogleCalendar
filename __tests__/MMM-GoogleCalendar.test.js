// Mock the global Module object and its register method
const mockModule = {
  register: jest.fn((name, moduleObject) => {
    // Store the registered module object so we can access its methods
    global.MMMGoogleCalendarInstance = moduleObject;
  })
};
global.Module = mockModule;

// Mock Log and moment, as they are used in start() and other places,
// though not directly by capFirst. This prevents errors if other parts of the module
// are inadvertently touched during instantiation.
global.Log = {
  info: jest.fn(),
  debug: jest.fn(),
  error: jest.fn(),
  warn: jest.fn()
};
global.moment = jest.fn(() => ({
  updateLocale: jest.fn(),
  startOf: jest.fn().mockReturnThis(),
  endOf: jest.fn().mockReturnThis(),
  clone: jest.fn().mockReturnThis(),
  add: jest.fn().mockReturnThis(),
  subtract: jest.fn().mockReturnThis(),
  toDate: jest.fn().mockReturnValue(new Date(0)),
  format: jest.fn().mockReturnValue("mocked_date"),
  valueOf: jest.fn().mockReturnValue(0), // for extractCalendarDate
  calendar: jest.fn().mockReturnValue("mocked_calendar_date"),
  fromNow: jest.fn().mockReturnValue("mocked_fromNow"),
  localeData: jest.fn(() => ({ // Mock for getLocaleSpecification
      longDateFormat: jest.fn().mockReturnValue("LT")
  }))
}));
global.config = { language: 'en', timeFormat: '24' }; // Mock global config

// Now, require the module. This will execute Module.register
require('../MMM-GoogleCalendar.js');

describe('MMM-GoogleCalendar', () => {
  let GCal;

  beforeEach(() => {
    // Access the stored module instance
    GCal = global.MMMGoogleCalendarInstance;
    // Ensure a clean state for config if needed by other tests, though capFirst is simple.
    GCal.config = { ...GCal.defaults, ...global.config }; // Reset config to defaults + global mock
  });

  describe('capFirst', () => {
    test('should capitalize the first letter of a lowercase string', () => {
      expect(GCal.capFirst('hello')).toBe('Hello');
    });

    test('should return an already capitalized string as is', () => {
      expect(GCal.capFirst('World')).toBe('World');
    });

    test('should return an empty string if input is empty', () => {
      expect(GCal.capFirst('')).toBe('');
    });

    test('should capitalize a single character string', () => {
      expect(GCal.capFirst('a')).toBe('A');
    });

    test('should handle strings with leading/trailing spaces (capFirst does not trim)', () => {
      expect(GCal.capFirst(' test ')).toBe(' test '); // Corrected expectation
    });

    test('should handle strings with numbers or symbols at the beginning', () => {
      expect(GCal.capFirst('1test')).toBe('1test');
      expect(GCal.capFirst('$test')).toBe('$test');
    });
  });

  describe('extractCalendarDate', () => {
    test('parses an all-day event (date)', () => {
      expect(GCal.extractCalendarDate({ date: '2026-06-05' })).toBe(0); // mocked valueOf
    });

    test('parses a timed event (dateTime)', () => {
      expect(GCal.extractCalendarDate({ dateTime: '2026-06-05T10:00:00Z' })).toBe(0);
    });

    test('returns null for undefined without throwing (issue #27)', () => {
      expect(() => GCal.extractCalendarDate(undefined)).not.toThrow();
      expect(GCal.extractCalendarDate(undefined)).toBeNull();
    });

    test('returns null for null without throwing', () => {
      expect(GCal.extractCalendarDate(null)).toBeNull();
    });
  });

  describe('createEventList', () => {
    test('skips events missing start/end instead of aborting the whole render (issue #27)', () => {
      // Only configured calendars are read (socketNotificationReceived
      // already drops data for unconfigured IDs).
      GCal.config.calendars = [{ calendarID: 'cal1' }];
      GCal.calendarData = {
        cal1: [
          { id: 'good', summary: 'Good', start: { dateTime: 'x' }, end: { dateTime: 'y' } },
          { id: 'cancelled', status: 'cancelled' }, // no start/end
          { id: 'good2', summary: 'Good2', start: { dateTime: 'x' }, end: { dateTime: 'y' } }
        ]
      };

      let result;
      expect(() => {
        result = GCal.createEventList();
      }).not.toThrow();

      // The malformed event is dropped, the valid ones survive.
      const ids = result.map((e) => e.id);
      expect(ids).toContain('good');
      expect(ids).toContain('good2');
      expect(ids).not.toContain('cancelled');
    });
  });

  describe('isEventExcluded', () => {
    test('matches a case-insensitive substring (issue #55)', () => {
      GCal.config.excludedEvents = ['Birthday'];
      expect(GCal.isEventExcluded("John's Birthday")).toBe(true);
      expect(GCal.isEventExcluded('my birthday party')).toBe(true);
    });

    test('does not match unrelated titles', () => {
      GCal.config.excludedEvents = ['Birthday'];
      expect(GCal.isEventExcluded('Team standup')).toBe(false);
    });

    test('still matches an exact full title (backwards compatible)', () => {
      GCal.config.excludedEvents = ['Lunch'];
      expect(GCal.isEventExcluded('Lunch')).toBe(true);
    });

    test('object form with caseSensitive only matches exact case', () => {
      GCal.config.excludedEvents = [{ filterBy: 'Standup', caseSensitive: true }];
      expect(GCal.isEventExcluded('Daily Standup')).toBe(true);
      expect(GCal.isEventExcluded('daily standup')).toBe(false);
    });

    test('object form with regex matches a pattern', () => {
      GCal.config.excludedEvents = [{ filterBy: '^\\[private\\]', regex: true }];
      expect(GCal.isEventExcluded('[private] dentist')).toBe(true);
      expect(GCal.isEventExcluded('not [private]')).toBe(false);
    });

    test('returns false for a private event with no summary', () => {
      GCal.config.excludedEvents = ['Birthday'];
      expect(GCal.isEventExcluded(undefined)).toBe(false);
    });

    test('ignores empty/invalid filter entries', () => {
      GCal.config.excludedEvents = ['', { filterBy: '' }, {}];
      expect(GCal.isEventExcluded('anything')).toBe(false);
    });
  });

  describe('fullDayEventRelativeLabel', () => {
    test('builds a date-only label from core TODAY/TOMORROW keys, no time (issue #85)', () => {
      GCal.translate = jest.fn((key) => key);
      GCal.config.dateFormat = 'MMM Do';

      const label = GCal.fullDayEventRelativeLabel(0);

      // Returns a non-empty capitalized string (moment is mocked).
      expect(typeof label).toBe('string');
      expect(label.length).toBeGreaterThan(0);

      // Uses date-only relative keys (no "at LT" time component).
      expect(GCal.translate).toHaveBeenCalledWith('TODAY');
      expect(GCal.translate).toHaveBeenCalledWith('TOMORROW');
    });
  });

  describe('colorForEvent (issue #75)', () => {
    beforeEach(() => {
      GCal.config.calendars = [{ calendarID: 'cal-a', color: '#abcdef' }];
    });

    test('falls back to the calendar color when the feature is off', () => {
      GCal.config.useGoogleEventColors = false;
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '11' })).toBe('#abcdef');
    });

    test('maps a Google colorId to the modern palette by default', () => {
      GCal.config.useGoogleEventColors = true;
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '11' })).toBe('#d50000');
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '1' })).toBe('#7986cb');
    });

    test('maps to the classic palette when asked for it', () => {
      GCal.config.useGoogleEventColors = true;
      GCal.config.googleEventColorPalette = 'classic';
      // The values colors.get actually returns.
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '11' })).toBe('#dc2127');
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '1' })).toBe('#a4bdfc');
      // Graphite is near-white in classic and dark grey in modern, so it is
      // the clearest signal that the right palette is in play.
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '8' })).toBe('#e1e1e1');
    });

    test('falls back to the modern palette for an unknown palette name', () => {
      GCal.config.useGoogleEventColors = true;
      GCal.config.googleEventColorPalette = 'neon';
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '8' })).toBe('#616161');
    });

    test('both palettes define all eleven Google color IDs', () => {
      for (const palette of ['modern', 'classic']) {
        const ids = Object.keys(GCal.googleEventColors[palette]);
        expect(ids).toHaveLength(11);
        expect(ids.sort((a, b) => a - b)).toEqual(
          ['1','2','3','4','5','6','7','8','9','10','11'].sort((a, b) => a - b)
        );
      }
    });

    test('accepts numeric colorIds as well as strings', () => {
      GCal.config.useGoogleEventColors = true;
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: 5 })).toBe('#f6bf26');
    });

    test('falls back to the calendar color for events with no colorId', () => {
      // Events using their calendar's default color carry no colorId at all.
      GCal.config.useGoogleEventColors = true;
      expect(GCal.colorForEvent({ calendarID: 'cal-a' })).toBe('#abcdef');
    });

    test('honours googleEventColorOverrides ahead of either palette', () => {
      GCal.config.useGoogleEventColors = true;
      GCal.config.googleEventColorOverrides = { 8: '#ffffff' };
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '8' })).toBe('#ffffff');
      // Unrelated IDs keep the built-in value.
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '2' })).toBe('#33b679');

      // An override wins over the selected palette, not just the default one.
      GCal.config.googleEventColorPalette = 'classic';
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '8' })).toBe('#ffffff');
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '2' })).toBe('#7ae7bf');
    });

    test('falls back to the calendar color for an unknown colorId', () => {
      GCal.config.useGoogleEventColors = true;
      expect(GCal.colorForEvent({ calendarID: 'cal-a', colorId: '99' })).toBe('#abcdef');
      expect(global.Log.warn).toHaveBeenCalled();
    });

    test('falls back to the default color for an unknown calendar', () => {
      GCal.config.useGoogleEventColors = true;
      expect(GCal.colorForEvent({ calendarID: 'nope' })).toBe('#fff');
    });
  });

  describe('mergeEventsById', () => {
    const A = 'a@example.com';
    const B = 'b@example.com';
    const C = 'c@example.com';

    const setup = (calendarData, accessRole = {}, order = [A, B, C]) => {
      GCal.config.calendars = order.map((calendarID) => ({ calendarID }));
      GCal.calendarData = calendarData;
      GCal.calendarAccessRole = accessRole;
    };

    test('keeps one copy per id and records every calendar in config order', () => {
      setup({
        [B]: [{ id: 'shared', summary: 'Daycare Tour' }],
        [A]: [{ id: 'shared', summary: 'Daycare Tour' }, { id: 'only-a' }]
      });
      const merged = GCal.mergeEventsById();
      expect(merged).toHaveLength(2);
      const shared = merged.find((e) => e.id === 'shared');
      expect(shared.calendarIDs).toEqual([A, B]);
      expect(merged.find((e) => e.id === 'only-a').calendarIDs).toEqual([A]);
    });

    test("readable attendee copy beats organizer's hidden copy", () => {
      setup(
        {
          [A]: [{ id: 'x', visibility: 'private', organizer: { self: true }, updated: '2026-11-06T15:40:43Z' }],
          [B]: [{ id: 'x', visibility: 'private', summary: 'Full details', organizer: {}, updated: '2026-11-06T15:40:40Z' }]
        },
        { [A]: 'reader', [B]: 'writer' }
      );
      const [event] = GCal.mergeEventsById();
      expect(event.calendarID).toBe(B);
      expect(event.summary).toBe('Full details');
      expect(event.calendarIDs).toEqual([A, B]);
    });

    test("organizer's copy wins even when listed later in config", () => {
      setup({
        [A]: [{ id: 'x', organizer: {}, updated: '2026-11-07T00:00:00Z' }],
        [B]: [{ id: 'x', organizer: { self: true }, updated: '2026-11-06T00:00:00Z' }]
      });
      expect(GCal.mergeEventsById()[0].calendarID).toBe(B);
    });

    test('without an organizer copy, the newest updated wins', () => {
      setup({
        [A]: [{ id: 'x', updated: '2026-11-06T00:00:00Z' }],
        [B]: [{ id: 'x', updated: '2026-11-07T00:00:00Z' }]
      });
      expect(GCal.mergeEventsById()[0].calendarID).toBe(B);
    });

    test('on a full tie, the first calendar in config order wins', () => {
      setup({
        [A]: [{ id: 'x', updated: '2026-11-06T00:00:00Z' }],
        [B]: [{ id: 'x', updated: '2026-11-06T00:00:00Z' }]
      }, {}, [B, A]);
      expect(GCal.mergeEventsById()[0].calendarID).toBe(B);
    });

    test('recurring instances are matched per instance id', () => {
      setup({
        [A]: [{ id: 'base_20261118T163000Z', recurringEventId: 'base' }, { id: 'base_20261125T163000Z', recurringEventId: 'base' }],
        [B]: [{ id: 'base_20261118T163000Z', recurringEventId: 'base' }]
      });
      const merged = GCal.mergeEventsById();
      expect(merged).toHaveLength(2);
      expect(merged.find((e) => e.id === 'base_20261118T163000Z').calendarIDs).toEqual([A, B]);
      expect(merged.find((e) => e.id === 'base_20261125T163000Z').calendarIDs).toEqual([A]);
    });

    test('events without an id are not merged', () => {
      setup({ [A]: [{ summary: 'no id' }], [B]: [{ summary: 'no id' }] });
      const merged = GCal.mergeEventsById();
      expect(merged).toHaveLength(2);
      expect(merged.map((e) => e.calendarIDs)).toEqual([[A], [B]]);
    });

    test('does not mutate the stored calendar data', () => {
      const raw = { id: 'x' };
      setup({ [A]: [raw] });
      GCal.mergeEventsById();
      expect(raw).toEqual({ id: 'x' });
    });
  });

  describe('isDetailsHidden', () => {
    const cal = 'c@example.com';
    const hidden = (role, visibility) => {
      GCal.calendarAccessRole = role === undefined ? {} : { [cal]: role };
      return GCal.isDetailsHidden({ visibility }, cal);
    };

    test('reader + private is hidden', () => expect(hidden('reader', 'private')).toBe(true));
    test('reader + confidential is hidden', () => expect(hidden('reader', 'confidential')).toBe(true));
    test('reader + default is readable', () => expect(hidden('reader', 'default')).toBe(false));
    test('writer + private is readable', () => expect(hidden('writer', 'private')).toBe(false));
    test('unknown role is readable', () => expect(hidden(undefined, 'private')).toBe(false));
  });
});
