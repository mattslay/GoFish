# How to contribute to GoFish

## Report a bug
- Please check [issues](https://github.com/VFPXGoFish/issues) if the bug has already been reported.
- If you're unable to find an open issue addressing the problem, open a new one. Be sure to include a title and clear description, as much relevant information as possible, and a code sample or an executable test case demonstrating the expected behavior that is not occurring.

## Fix a bug or add an enhancement
- Fork the project: see this [guide](https://www.dataschool.io/how-to-contribute-on-github/) for setting up and using a fork.
- Make whatever changes are necessary.
- Add a description of the changes to `Source\changelog_ver_5.txt`. This is:
   - *Version number* and *Release date*
   - Changes for the Version
- Copy the added text *Changes for the Version* from `Source\changelog_ver_5.txt` to the *Release History section* of `readme.md` and change, at the top of the document, the
  - version number and
  - date .
- In `Source\BuildGoFish.prg`, change at top the 
  - *lcVersion* and 
  - *lcBuild* variables   

- run `Source\BuildGoFish.prg` to
  - update several files
  - build GoFish5.app,
   - and generate the text equivalents for all VFP binary files (SCX, VCX, DBF, etc.) using FoxBin2PRG.
   > Note: use VFP 9 SP2 rather than VFP Advanced to run BuildGoFish.prg since GoFish5.app must be built in VFP rather than VFPA for compatibility for all users.

- In File Explorer, right-click `Source\BuildCloudZip.ps1` and choose *Run with PowerShell* to create Source.zip.
- Commit the changes.
- Push to your fork.
- Create a pull request; ensure the description clearly describes the problem and solution or the enhancement.

----
Last changed: 2022-10-16