"""
MimicMail — Multi-Instance Parallel Orchestrator & Launcher
Run multiple independent MimicMail instances in parallel with isolated profiles.
"""

import sys
import os
import subprocess
import argparse
import time

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MAIN_SCRIPT = os.path.join(SCRIPT_DIR, "zoho_playwright_mailer.py")

def launch_parallel_instances(count: int = 2, start_index: int = 1):
    python_exe = sys.executable
    processes = []

    print("=" * 65)
    print(f"🚀 MimicMail Parallel Orchestrator")
    print(f"Launching {count} isolated, simultaneous instance(s)...")
    print("=" * 65)

    for i in range(count):
        inst_num = start_index + i
        inst_name = f"Instance {inst_num}"
        print(f"  • Starting {inst_name} (Profile: instance_{inst_num})...")

        # Launch process independently without blocking the terminal
        p = subprocess.Popen([python_exe, MAIN_SCRIPT, "--instance", inst_name])
        processes.append((inst_name, p))
        time.sleep(1.0)  # Smooth stagger to prevent window overlap

    print("-" * 65)
    print(f"✓ All {count} instance(s) are active and running in parallel.")
    print("Each instance has its own separate cookies, Zoho session, and window.")
    print("=" * 65)

    return processes

def main():
    parser = argparse.ArgumentParser(description="Launch multiple MimicMail instances in parallel.")
    parser.add_argument(
        "-n", "--count",
        type=int,
        default=2,
        help="Number of parallel instances to launch simultaneously (default: 2)"
    )
    parser.add_argument(
        "-s", "--start",
        type=int,
        default=1,
        help="Starting instance number (default: 1)"
    )
    args = parser.parse_args()

    launch_parallel_instances(count=max(1, args.count), start_index=max(1, args.start))

if __name__ == "__main__":
    main()
