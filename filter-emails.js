#!/usr/bin/env node
'use strict';
/**
 * Filter recent emails by date.
 *
 * Security:
 *   - Honors --account=<name> like the other scripts (no silent cross-account leak).
 *   - Sanitizes every printed field to prevent terminal escape injection.
 *   - Validates the cutoff date.
 */

const { getRecent } = require('./email.js');
const { parseAccountFlag, sanitizeForTerminal, scrubSecrets } = require('./security.js');

(async () => {
  let accountName;
  try {
    accountName = parseAccountFlag(process.argv);
  } catch (e) {
    console.error('❌', e.message);
    process.exit(1);
  }

  const positional = process.argv.slice(2).filter((a) => !a.startsWith('--account='));
  const sinceArg = positional[0] || new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();
  const cutoff = new Date(sinceArg);
  if (isNaN(cutoff.getTime())) {
    console.error('❌ Invalid date:', sanitizeForTerminal(sinceArg));
    process.exit(1);
  }

  try {
    const emails = await getRecent(30, accountName);
    const recentEmails = emails.filter((e) => new Date(e.receivedDateTime) >= cutoff);

    console.log(`Emails since ${cutoff.toLocaleString('en-US')}:\n`);

    recentEmails.forEach((e) => {
      const date = new Date(e.receivedDateTime);
      const isUnread = e.isRead === false ? '🔵 ' : '';
      const from = sanitizeForTerminal((e.from && e.from.emailAddress && (e.from.emailAddress.name || e.from.emailAddress.address)) || 'Unknown');
      const subject = sanitizeForTerminal(e.subject || '(no subject)');

      console.log(`${isUnread}${subject}`);
      console.log(`  From: ${from}`);
      console.log(`  Date: ${date.toLocaleString('en-US')}`);

      if (e.bodyPreview) {
        console.log(`  Preview: ${sanitizeForTerminal(e.bodyPreview.substring(0, 100))}...`);
      }
      console.log('');
    });

    console.log(`Total: ${recentEmails.length} emails`);
  } catch (err) {
    console.error('❌ Error:', scrubSecrets(err.message));
    process.exit(1);
  }
})();
