#!/usr/bin/env node
'use strict';
/**
 * Microsoft Graph Email Operations
 *
 * Security:
 *   - All Graph IDs are validated before being concatenated into URLs (no path traversal).
 *   - Search queries are URL-encoded *and* stripped of `"` to prevent OData injection.
 *   - All terminal output (subjects, sender names, previews) is run through an
 *     ANSI/C0 sanitizer so a malicious email body cannot hijack the terminal.
 *   - HTTPS calls go through the shared, host-pinned, timeout-bounded helper.
 */

const { getAccessToken } = require('./auth.js');
const {
  assertGraphId,
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

function sanitizeSearchQuery(q) {
  if (typeof q !== 'string') throw new Error('Search query must be a string');
  if (q.length > 1000) throw new Error('Search query too long');
  // Strip control chars and double-quotes (we wrap the value in quotes ourselves).
  const cleaned = q.replace(/[\x00-\x1f"]/g, ' ').trim();
  if (!cleaned) throw new Error('Empty search query');
  return cleaned;
}

async function searchEmails(query, top = 10, accountName = null) {
  const safeQuery = sanitizeSearchQuery(query);
  const safeTop = Math.max(1, Math.min(100, parseInt(top, 10) || 10));
  const token = await getAccessToken(accountName);
  const url = `/v1.0/me/messages?$search="${encodeURIComponent(safeQuery)}"&$top=${safeTop}&$orderby=receivedDateTime desc`;
  const response = await graphGet(url, token);
  return response.value || [];
}

async function getFromSender(senderEmail, top = 10, accountName = null) {
  const safeSender = sanitizeSearchQuery(senderEmail);
  const safeTop = Math.max(1, Math.min(100, parseInt(top, 10) || 10));
  const token = await getAccessToken(accountName);
  const url = `/v1.0/me/messages?$search="from:${encodeURIComponent(safeSender)}"&$top=${safeTop}&$orderby=receivedDateTime desc`;
  const response = await graphGet(url, token);
  return response.value || [];
}

async function getRecent(top = 10, accountName = null) {
  const safeTop = Math.max(1, Math.min(100, parseInt(top, 10) || 10));
  const token = await getAccessToken(accountName);
  const url = `/v1.0/me/messages?$top=${safeTop}&$orderby=receivedDateTime desc`;
  const response = await graphGet(url, token);
  return response.value || [];
}

async function getEmailById(emailId, accountName = null) {
  assertGraphId(emailId, 'emailId');
  const token = await getAccessToken(accountName);
  const url = `/v1.0/me/messages/${encodeURIComponent(emailId)}`;
  return await graphGet(url, token);
}

function formatEmail(email, includeBody = false) {
  const date = new Date(email.receivedDateTime);
  const fromRaw = (email.from && email.from.emailAddress && (email.from.emailAddress.name || email.from.emailAddress.address)) || 'Unknown';
  const from = sanitizeForTerminal(fromRaw);
  const subject = sanitizeForTerminal(email.subject || '(no subject)');

  let output = `📧 ${subject}\n`;
  output += `   From: ${from}\n`;
  output += `   Date: ${date.toLocaleString('en-US')}\n`;

  if (email.isRead === false) {
    output += `   Status: 🔵 Unread\n`;
  }

  if (includeBody) {
    const body = (email.body && email.body.content) || email.bodyPreview || '';
    const plainText = body.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').trim();
    output += `\n${sanitizeForTerminal(plainText)}\n`;
  } else if (email.bodyPreview) {
    output += `   Preview: ${sanitizeForTerminal(email.bodyPreview.substring(0, 150))}...\n`;
  }

  return output;
}

if (require.main === module) {
  (async () => {
    const command = process.argv[2];
    const arg = process.argv[3];
    let accountName;
    try {
      accountName = parseAccountFlag(process.argv);
    } catch (e) {
      console.error('❌', e.message);
      process.exit(1);
    }

    try {
      if (command === 'search' && arg) {
        const emails = await searchEmails(arg, 10, accountName);
        if (emails.length === 0) return console.log('No emails found');
        console.log(`Found ${emails.length} email(s):\n`);
        emails.forEach((email) => console.log(formatEmail(email)));
      } else if (command === 'from' && arg) {
        const emails = await getFromSender(arg, 10, accountName);
        if (emails.length === 0) return console.log('No emails found from that sender');
        console.log(`Found ${emails.length} email(s) from ${sanitizeForTerminal(arg)}:\n`);
        emails.forEach((email) => console.log(formatEmail(email)));
      } else if (command === 'recent') {
        const count = arg ? parseInt(arg, 10) : 10;
        const emails = await getRecent(count, accountName);
        if (emails.length === 0) return console.log('No emails found');
        console.log(`${emails.length} most recent email(s):\n`);
        emails.forEach((email) => console.log(formatEmail(email)));
      } else if (command === 'read' && arg) {
        const email = await getEmailById(arg, accountName);
        console.log(formatEmail(email, true));
      } else {
        console.log('Usage:');
        console.log('  node email.js search "query" [--account=name]       - Search emails');
        console.log('  node email.js from email@domain [--account=name]    - Get emails from sender');
        console.log('  node email.js recent [count] [--account=name]       - Get recent emails');
        console.log('  node email.js read <id> [--account=name]            - Read full email');
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
  searchEmails,
  getFromSender,
  getRecent,
  getEmailById,
  formatEmail
};
