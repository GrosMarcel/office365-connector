#!/usr/bin/env node
'use strict';
/**
 * Microsoft Graph - Cancel Calendar Event
 *
 * Security:
 *   - eventId is validated before being placed in the URL path.
 *   - Comment is length-bounded.
 *   - HTTPS request is host-pinned and timeout-bounded.
 */

const { getAccessToken } = require('./auth.js');
const {
  assertGraphId,
  httpsJsonRequest,
  parseAccountFlag,
  scrubSecrets
} = require('./security.js');

const GRAPH_BASE = 'https://graph.microsoft.com';

async function cancelEvent(eventId, comment = '', accountName = null) {
  assertGraphId(eventId, 'eventId');
  if (typeof comment !== 'string' || comment.length > 4000) {
    throw new Error('Invalid comment (string, max 4000 chars)');
  }

  const token = await getAccessToken(accountName);
  const url = `${GRAPH_BASE}/v1.0/me/events/${encodeURIComponent(eventId)}/cancel`;
  const body = JSON.stringify({ comment });

  return await httpsJsonRequest(url, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(body)
    },
    body,
    acceptEmpty: true
  });
}

if (require.main === module) {
  (async () => {
    const args = process.argv.slice(2);
    let accountName;
    try {
      accountName = parseAccountFlag(process.argv);
    } catch (e) {
      console.error('❌', e.message);
      process.exit(1);
    }
    const filteredArgs = args.filter((arg) => !arg.startsWith('--account='));

    const eventId = filteredArgs[0];
    const comment = filteredArgs[1] || '';

    if (!eventId) {
      console.log('Usage: node cancel-event.js <event-id> [comment] [--account=name]');
      console.log('\nIf --account is not specified, the default account is used.');
      process.exit(1);
    }

    try {
      await cancelEvent(eventId, comment, accountName);
      console.log('✅ Event cancelled successfully');
    } catch (err) {
      console.error('❌ Error:', scrubSecrets(err.message));
      process.exit(1);
    }
  })();
}

module.exports = {
  cancelEvent
};
