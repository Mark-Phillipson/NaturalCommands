# Mouse movement commands for NaturalCommands
# Copy these into your Talon user directory

# Basic directional movement
^mouse left$:
    user.run_application_csharp_natural("mouse left")
^mouse right$:
    user.run_application_csharp_natural("mouse right")
^mouse up$:
    user.run_application_csharp_natural("mouse up")
^mouse down$:
    user.run_application_csharp_natural("mouse down")

# Diagonal movement
^mouse up left$:
    user.run_application_csharp_natural("mouse up left")
^mouse up right$:
    user.run_application_csharp_natural("mouse up right")
^mouse down left$:
    user.run_application_csharp_natural("mouse down left")
^mouse down right$:
    user.run_application_csharp_natural("mouse down right")

# Stop commands
^mouse stop$:
    user.run_application_csharp_natural("mouse stop")
^stop mouse$:
    user.run_application_csharp_natural("stop mouse")
^stop click$:
    user.run_application_csharp_natural("stop click")
^stop right click$:
    user.run_application_csharp_natural("stop right click")

# Speed adjustment
^mouse faster$:
    user.run_application_csharp_natural("mouse faster")
^mouse slower$:
    user.run_application_csharp_natural("mouse slower")
^faster$:
    user.run_application_csharp_natural("faster")
^slower$:
    user.run_application_csharp_natural("slower")
