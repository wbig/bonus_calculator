from summarizing import ctrl_panel

try:
    ctrl_panel()
except KeyboardInterrupt:
    print("\nThe program has been interrupted by the user.")