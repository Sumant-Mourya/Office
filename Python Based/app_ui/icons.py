"""Premium SVG icon definitions for the ActivityTracker UI."""

# Sleek Logo: Glowing stylized clock with gears and gradient pulse ring
TRACKER_LOGO = """
<svg width="48" height="48" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="drop-shadow-[0_0_8px_rgba(99,102,241,0.5)]">
    <circle cx="12" cy="12" r="10" stroke="url(#logo-grad)" stroke-width="2" stroke-linecap="round"/>
    <path d="M12 7V12L15 13.5" stroke="url(#logo-grad)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
    <circle cx="12" cy="12" r="2" fill="#818cf8"/>
    <defs>
        <linearGradient id="logo-grad" x1="2" y1="2" x2="22" y2="22" gradientUnits="userSpaceOnUse">
            <stop stop-color="#a78bfa"/>
            <stop offset="0.5" stop-color="#6366f1"/>
            <stop offset="1" stop-color="#3b82f6"/>
        </linearGradient>
    </defs>
</svg>
"""

# Header Title Icon
CONSOLE_HEADER_ICON = """
<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="mr-2">
    <path d="M12 2C6.48 2 2 6.48 2 12C2 17.52 6.48 22 12 22C17.52 22 22 17.52 22 12C22 6.48 17.52 2 12 2ZM13 17H11V15H13V17ZM13 13H11V7H13V13Z" fill="url(#header-grad)"/>
    <defs>
        <linearGradient id="header-grad" x1="2" y1="2" x2="22" y2="22" gradientUnits="userSpaceOnUse">
            <stop stop-color="#818cf8"/>
            <stop offset="1" stop-color="#6366f1"/>
        </linearGradient>
    </defs>
</svg>
"""

# Google Sheets Green Grid Icon
SHEET_ICON = """
<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M19 3H5C3.9 3 3 3.9 3 5V19C3 20.1 3.9 21 5 21H19C20.1 21 21 20.1 21 19V5C21 3.9 20.1 3 19 3ZM10 17H5V13H10V17ZM10 11H5V7H10V11ZM19 17H12V13H19V17ZM19 11H12V7H19V11Z" fill="#10b981"/>
</svg>
"""

# Modern computer monitor
COMPUTER_ICON = """
<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M21 2H3C1.9 2 1 2.9 1 4V16C1 17.1 1.9 18 3 18H10V21H8V23H16V21H14V18H21C22.1 18 23 17.1 23 16V4C23 2.9 22.1 2 21 2ZM21 14H3V4H21V14Z" fill="#6366f1"/>
</svg>
"""

# Floppy Diskette save icon
SAVE_ICON = """
<svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="mr-2 inline">
    <path d="M17 3H5C3.89 3 3 3.9 3 5V19C3 20.1 3.89 21 5 21H19C20.1 21 21 20.1 21 19V7L17 3ZM12 19C10.34 19 9 17.66 9 16C9 14.34 10.34 13 12 13C13.66 13 15 14.34 15 16C15 17.66 13.66 19 12 19ZM15 9H5V5H15V9Z" fill="currentColor"/>
</svg>
"""

# Play arrow tracking start icon
START_ICON = """
<svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="mr-2 inline">
    <path d="M8 5V19L19 12L8 5Z" fill="currentColor"/>
</svg>
"""

# Stop/Unlock tracking stop icon
STOP_ICON = """
<svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="mr-2 inline">
    <path d="M18 18H6V6H18V18Z" fill="currentColor"/>
</svg>
"""

# Refresh/Sync icon
REFRESH_ICON = """
<svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="mr-1 inline">
    <path d="M17.65 6.35C16.2 4.9 14.21 4 12 4C7.58 4 4.01 7.58 4.01 12C4.01 16.42 7.58 20 12 20C15.73 20 18.84 17.45 19.73 14H17.65C16.83 16.33 14.61 18 12 18C8.69 18 6 15.31 6 12C6 8.69 8.69 6 12 6C13.66 6 15.14 6.69 16.22 7.78L13 11H20V4L17.65 6.35Z" fill="currentColor"/>
</svg>
"""

# User silhouette icon for login page
USER_ICON = """
<svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="mr-2 inline">
    <path d="M12 12C14.21 12 16 10.21 16 8C16 5.79 14.21 4 12 4C9.79 4 8 5.79 8 8C8 10.21 9.79 12 12 12ZM12 14C9.33 14 4 15.33 4 18V20H20V18C20 15.33 14.67 14 12 14Z" fill="currentColor"/>
</svg>
"""

# Lock padlock icon for login page
LOCK_ICON = """
<svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="mr-2 inline">
    <path d="M18 8H17V6C17 3.24 14.76 1 12 1C9.24 1 7 3.24 7 6V8H6C4.9 8 4 8.9 4 10V20C4 21.1 4.9 22 6 22H18C19.1 22 20 21.1 20 20V10C20 8.9 19.1 8 18 8ZM9 6C9 4.34 10.34 3 12 3C13.66 3 15 4.34 15 6V8H9V6ZM18 20H6V10H18V20ZM12 17C13.1 17 14 16.1 14 15C14 13.9 13.1 13 12 13C10.9 13 10 13.9 10 15C10 16.1 10.9 17 12 17Z" fill="currentColor"/>
</svg>
"""

# Interactive Keyboard Icon
KEYBOARD_ICON = """
<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="inline-block mr-1">
    <path d="M20 5H4C2.9 5 2 5.9 2 7V17C2 18.1 2.9 19 4 19H20C21.1 19 22 18.1 22 17V7C22 5.9 21.1 5 20 5ZM20 17H4V7H20V17ZM5 8H7V10H5V8ZM5 11H7V13H5V11ZM5 14H7V16H5V14ZM8 8H10V10H8V8ZM8 11H10V13H8V11ZM8 14H10V16H8V14ZM11 8H13V10H11V8ZM11 11H13V13H11V11ZM11 14H16V16H11V14ZM14 8H16V10H14V8ZM14 11H16V13H14V11ZM17 8H19V10H17V8ZM17 11H19V13H17V11ZM17 14H19V16H17V14Z" fill="currentColor"/>
</svg>
"""

# Interactive Mouse Icon
MOUSE_ICON = """
<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="inline-block mr-1">
    <path d="M12 2C8.69 2 6 4.69 6 8V16C6 19.31 8.69 22 12 22C15.31 22 18 19.31 18 16V8C18 4.69 15.31 2 12 2ZM16 16C16 18.21 14.21 20 12 20C9.79 20 8 18.21 8 16V11H16V16ZM16 9H8V8C8 5.79 9.79 4 12 4C14.21 4 16 5.79 16 8V9Z" fill="currentColor"/>
</svg>
"""

# Aggregated clock outline icon
CLOCK_ICON = """
<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" class="inline-block mr-1">
    <path d="M11.99 2C6.47 2 2 6.48 2 12C2 17.52 6.47 22 11.99 22C17.52 22 22 17.52 22 12C22 6.48 17.52 2 11.99 2ZM12 20C7.58 20 4 16.42 4 12C4 7.58 7.58 4 12 4C16.42 4 20 7.58 20 12C20 16.42 16.42 20 12 20ZM12.5 7H11V13L16.25 16.15L17 14.92L12.5 12.25V7Z" fill="currentColor"/>
</svg>
"""
