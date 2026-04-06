from plyer import notification

print("\n  Sending test toast notification...")
notification.notify(
    title="US Economic Tracker",
    message="REGIME CHANGE: Reflation -> Stagflation (94% confidence)",
    timeout=10,
)
print("  Toast sent - check the bottom-right of your screen\n")
