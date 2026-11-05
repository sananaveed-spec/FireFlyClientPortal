#!/bin/sh
set -e
# Named volume may be root-owned from first create; app runs as nextjs (1001).
if [ -d /app/data ]; then
  chown -R nextjs:nodejs /app/data 2>/dev/null || true
fi
exec su-exec nextjs "$@"
