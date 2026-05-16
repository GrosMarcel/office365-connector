#!/usr/bin/env node
'use strict';
/**
 * Microsoft Graph Calendar Operations
 *
 * Security:
 *   - HTTPS calls go through the shared host-pinned, timeout-bounded helper.
 *   - Event subjects, locations, organizer/attendee names are sanitized before
 *     being printed (no terminal escape injection from a hostile invite).
 */

const { getAccessToken } = require('./auth.js');
const {
  httpsJsonRequest,
  parseAccountFlag,
  sanitizeForTerminal,
  scrubSecrets
} = require('./security.js');

const GRAPH_BASE = 'https://graph.microsoft.com';

function graphGet(pathAndQuery, accessToken) {
  return httpsJsonRequest(GRAPH_BASE + pathAndQuery, {
    method: 'GET',
    headers: { 'Authorization': `Bearer ${accessToken}` }
  });
}

async function getEvents(startDate, endDate, accountName = null) {
  if (!(startDate instanceof Date) || isNaN(startDate.getTime())) throw new Error('Invalid startDate');
  if (!(endDate instanceof Date) || isNaN(endDate.getTime())) throw new Error('Invalid endDate');

  const token = await getAccessToken(accountName);

  const startISO = startDate.toISOString();
  const endISO = endDate.toISOString();

  const url = `/v1.0/me/calendarview?startDateTime=${encodeURIComponent(startISO)}&endDateTime=${encodeURIComponent(endISO)}&$orderby=start/dateTime&$top=50`;

  const response = await graphGet(url, token);
  return response.value || [];
}

function formatEvent(event) {
  const start = new Date(event.start.dateTime + 'Z');
  const end = new Date(event.end.dateTime + 'Z');

  const timeStr = event.isAllDay
    ? 'All Day'
    : `${start.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })} - ${end.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}`;

  let output = `${timeStr} | ${sanitizeForTerminal(event.subject || '(no subject)')}`;

  if (event.location && event.location.displayName) {
    output += ` | 📍 ${sanitizeForTerminal(event.location.displayName)}`;
  }

  if (event.organizer && event.organizer.emailAddress && event.organizer.emailAddress.name) {
    output += ` | 👤 ${sanitizeForTerminal(event.organizer.emailAddress.name)}`;
  }

  if (event.attendees && event.attendees.length > 0) {
    const organizerAddress = event.organizer && event.organizer.emailAddress && event.organizer.emailAddress.address;
    const attendeeNames = event.attendees
      .filter((a) => a.emailAddress.address !== organizerAddress)
      .map((a) => sanitizeForTerminal(a.emailAddress.name))
      .slice(0, 3);
    if (attendeeNames.length > 0) {
      output += ` | +${attendeeNames.join(', ')}`;
    }
  }

  return output;
}

async function getToday(accountName = null) {
  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  const endOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);

  const events = await getEvents(startOfDay, endOfDay, accountName);

  if (events.length === 0) {
    console.log('📅 No events scheduled for today');
    return;
  }

  console.log(`📅 Today's Calendar (${now.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' })})\n`);
  events.forEach((event) => console.log(formatEvent(event)));
  console.log(`\nTotal: ${events.length} event${events.length !== 1 ? 's' : ''}`);
}

async function getWeek(accountName = null) {
  const now = new Date();
  const startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay());
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(endOfWeek.getDate() + 7);

  const events = await getEvents(startOfWeek, endOfWeek, accountName);

  if (events.length === 0) {
    console.log('📅 No events scheduled this week');
    return;
  }

  console.log(`📅 This Week's Calendar\n`);

  let currentDay = null;
  events.forEach((event) => {
    const eventDate = new Date(event.start.dateTime + 'Z');
    const dayStr = eventDate.toLocaleDateString('en-US', { weekday: 'long', month: 'short', day: 'numeric' });

    if (dayStr !== currentDay) {
      if (currentDay !== null) console.log('');
      console.log(`\n${dayStr}:`);
      currentDay = dayStr;
    }

    console.log('  ' + formatEvent(event));
  });

  console.log(`\nTotal: ${events.length} event${events.length !== 1 ? 's' : ''}`);
}

if (require.main === module) {
  (async () => {
    const command = process.argv[2] || 'today';
    let accountName;
    try {
      accountName = parseAccountFlag(process.argv);
    } catch (e) {
      console.error('❌', e.message);
      process.exit(1);
    }

    try {
      if (command === 'today') {
        await getToday(accountName);
      } else if (command === 'week') {
        await getWeek(accountName);
      } else {
        console.log('Usage:');
        console.log('  node calendar.js today [--account=name]  - Show today\'s events');
        console.log('  node calendar.js week [--account=name]   - Show this week\'s events');
        console.log('\nIf --account is not specified, the default account is used.');
        process.exit(1);
      }
    } catch (err) {
      console.error('❌ Error:', scrubSecrets(err.message));
      process.exit(1);
    }
  })();
}

module.exports = {
  getEvents,
  getToday,
  getWeek,
  formatEvent
};
