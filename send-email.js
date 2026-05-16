#!/usr/bin/env node
'use strict';
/**
 * Microsoft Graph - Send Email
 *
 * Security:
 *   - Recipient addresses are validated before being placed in the JSON payload.
 *   - messageId is validated before being concatenated into the reply URL.
 *   - Outbound HTTPS calls are host-pinned, timed-out, and size-capped.
 *   - Error messages are scrubbed of any access_token / refresh_token leakage.
 */

const { getAccessToken } = require('./auth.js');
const {
  assertEmail,
  assertGraphId,
  httpsJsonRequest,
  parseAccountFlag,
  scrubSecrets
} = require('./security.js');

const GRAPH_BASE = 'https://graph.microsoft.com';

function graphPost(pathAndQuery, accessToken, jsonBody) {
  const body = JSON.stringify(jsonBody);
  return httpsJsonRequest(GRAPH_BASE + pathAndQuery, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(body)
    },
    body,
    acceptEmpty: true
  });
}

function toRecipientList(input, label) {
  const list = Array.isArray(input) ? input : [input];
  if (list.length === 0) throw new Error(`${label}: at least one recipient required`);
  if (list.length > 500) throw new Error(`${label}: too many recipients (max 500)`);
  return list.map((address) => ({ emailAddress: { address: assertEmail(address, label) } }));
}

async function sendEmail(to, subject, body, options = {}) {
  if (typeof subject !== 'string' || subject.length > 998) throw new Error('Invalid subject');
  if (typeof body !== 'string') throw new Error('Invalid body');
  if (body.length > 10 * 1024 * 1024) throw new Error('Body too large (>10 MB)');

  const token = await getAccessToken(options.accountName);

  const message = {
    message: {
      subject,
      body: {
        contentType: options.html ? 'HTML' : 'Text',
        content: body
      },
      toRecipients: toRecipientList(to, 'to')
    }
  };

  if (options.cc) message.message.ccRecipients = toRecipientList(options.cc, 'cc');
  if (options.replyTo) {
    message.message.replyTo = [{ emailAddress: { address: assertEmail(options.replyTo, 'replyTo') } }];
  }

  return await graphPost('/v1.0/me/sendMail', token, message);
}

async function replyToEmail(messageId, body, options = {}) {
  assertGraphId(messageId, 'messageId');
  if (typeof body !== 'string') throw new Error('Invalid body');
  if (body.length > 10 * 1024 * 1024) throw new Error('Body too large (>10 MB)');

  const token = await getAccessToken(options.accountName);
  const url = `/v1.0/me/messages/${encodeURIComponent(messageId)}/reply`;
  return await graphPost(url, token, { comment: body });
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

    try {
      if (filteredArgs[0] === 'send' && filteredArgs.length >= 3) {
        await sendEmail(filteredArgs[1], filteredArgs[2], filteredArgs[3] || '', { accountName });
        console.log('✅ Email sent successfully');
      } else if (filteredArgs[0] === 'reply' && filteredArgs.length >= 2) {
        await replyToEmail(filteredArgs[1], filteredArgs[2] || '', { accountName });
        console.log('✅ Reply sent successfully');
      } else {
        console.log('Usage:');
        console.log('  node send-email.js send <to> <subject> <body> [--account=name]');
        console.log('  node send-email.js reply <message-id> <body> [--account=name]');
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
  sendEmail,
  replyToEmail
};
