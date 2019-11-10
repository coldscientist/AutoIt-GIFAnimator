# Change Log

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## V1.0.1.0, 2019-09-11
### Added
- The script detects if GIFAnimator.exe process was closed and exits the script.
- The script uses GDI functions to detect GIF frame duration instead of relying on Microsoft GIF Animator, so the user can continue using the PC and will be interrupted only if the GIF frame duration must be changed.

### Changed
- You can close the script pressing `Ctrl` + `E` instead of `Esc`, avoiding closing the program by mistake.

### Fixed
- The script uses a "Timer" feature to detect if "Open" and "Save" dialog was opened with success and retry to open it after 5 seconds if the dialog wasn't detected.
- The script can detect if all frames was selected in Microsoft GIF Animator window and retry selecting all of them if it fails.
- Script could fail to detect "Ready" status after saving a file through Microsoft GIF Animator.

## V1.0.0.0, 2019-09-15
- First public release
