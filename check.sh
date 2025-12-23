# 1. Check cgroup v2 (most likely location)
cat /sys/fs/cgroup/memory.max /sys/fs/cgroup/memory.current 2>/dev/null | head -20

# 2. Check cgroup v1 (older Docker)
cat /sys/fs/cgroup/memory/memory.limit_in_bytes /sys/fs/cgroup/memory/memory.usage_in_bytes 2>/dev/null

# 3. Find ALL memory-related cgroup files
find /sys/fs/cgroup -name "*memory*" -type f 2>/dev/null | head -20

# 4. Check /proc/meminfo (this shows host, but useful to compare)
cat /proc/meminfo | grep -E "MemTotal|MemAvailable|MemFree"

# 5. Check what cgroup version you're using
cat /proc/self/cgroup

# 6. Check for any Docker environment variables
env | grep -i "mem\|limit"

# 7. Try to get memory from systemd if available
cat /proc/1/cgroup 2>/dev/null


# Check the actual memory cgroup path for this process
cat /sys/fs/cgroup/memory/user.slice/user-800540758.slice/session-c32.scope/memory.limit_in_bytes 2>/dev/null

# Check parent cgroup limits
find /sys/fs/cgroup/memory/user.slice -name "memory.limit_in_bytes" -exec sh -c 'echo "{}:" && cat {}' \; 2>/dev/null | grep -B1 -v 9223372036854771712

# Check systemd slice limits
systemctl show user-800540758.slice | grep Memory

# Or check via podman
podman stats --no-stream 2>/dev/null
