// Options for the DocFX "modern" template. The template imports this file and reads its default
// export, so it replaces the template's empty one.

export default {
  // The site is light only. This is the template's own option for the starting theme.
  defaultTheme: 'light',

  // defaultTheme alone is not enough: the template stores the reader's choice in localStorage and
  // that choice wins. A visitor who picked "dark" before this change, and the template's "auto"
  // (follow the operating system, which a dark-mode laptop would turn dark), must both end up on
  // light. start() runs before the template applies the theme, so writing the value here is enough.
  start () {
    try {
      localStorage.setItem('theme', 'light')
    } catch {
      // Storage can be blocked (private mode, strict privacy settings). The page then falls back to
      // defaultTheme above, which is still light.
    }
  }
}
