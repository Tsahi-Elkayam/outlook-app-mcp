# Security

This server gives an AI model access to your mail, calendar and contacts. Read
"Mail is untrusted input" and "Safety design" in the README before you use it,
and keep the tools that change or send things on "ask" in your client.

## Reporting a vulnerability

Please report security problems privately through GitHub's private
vulnerability reporting (the repository's Security tab > Report a
vulnerability), not in a public issue. Include the Outlook version, the
operating system and the steps that show the problem. You can expect a first
answer within a week.

Of most interest:

- a way to make a tool delete permanently, send without a sending tool, or
  write outside the allowed folders;
- a way for text in a message or event to reach Outlook as script or code
  instead of data;
- a read tool that changes anything.
