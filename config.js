/* ═══════════════════════════════════════════════════════════════
   TEAM CONFIG — this is the only file that changes per deployment
   Copy this file for each team and update the values below.
   dashboard.js and dashboard.css stay identical across all teams.
═══════════════════════════════════════════════════════════════ */

/* Microsoft identity */
var CLIENT_ID      = '4ae73f12-9d89-467d-9dec-79be71a18224';
var TENANT_ID      = '21827f82-8236-44d6-b5fe-bd2d5d65b3a9';
var REDIRECT       = 'https://shantala-mitchell.github.io/ITServiceDesk/dashboard.html';

/* SharePoint — update these for each team's list */
var SITE_ID        = 'cd2c6dbd-3685-41c9-9d76-820fb6f2350d';
var LIST_ID        = 'c83d54e5-0c6a-4274-aeb2-b584e48826b8';
var AGENTS_LIST_ID = 'c3b2a174-0cbc-4f8b-9d0f-075c60746aca';

/* Display name shown in the browser tab and reports */
var TEAM_NAME      = 'IT Service Desk';

/* Archive threshold — tickets closed longer than this (in days) move to Archive.
   This value is shared across all agents on this deployment.
   0 = archive immediately on close. */
var ARCHIVE_DAYS   = 0;

/* Fallback admin if the Agents list is empty or unreachable */
var DEFAULT_AGENTS = [{email: 'shantala.mitchell@ecnz.ac.nz', role: 'admin'}];
