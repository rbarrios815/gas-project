










README HERE

## Troubleshooting: legacy JB/RB text format

If you receive a JB/RB text with the old subject/body format (for example,
`DASHBOARD CLIENT TASKS FOR TODAY` with dated note lines), the usual cause is
another Apps Script deployment or copied project still running the legacy
code. Check:

- **Deployments**: confirm only the latest deployment is active in the Apps
  Script project.
- **Triggers**: remove any time-based triggers tied to older script versions
  or copies.
- **Copies**: search your Drive for copies of this script and verify they are
  not scheduled to send messages.
