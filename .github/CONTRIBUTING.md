# How to contribute to GoFish

## Report a bug
- Please check [issues](https://github.com/VFPXGoFish/issues) if the bug has already been reported.
- If you're unable to find an open issue addressing the problem, open a new one. Be sure to include a title and clear description, as much relevant information as possible, and a code sample or an executable test case demonstrating the expected behavior that is not occurring.

## Fix a bug or add an enhancement
- Fork the project: see this [guide](https://www.dataschool.io/how-to-contribute-on-github/) for setting up and using a fork.
  - If already forked, pull the recent state, or get most recent version otherwise.
- Make whatever changes are necessary.
- In **[docs\ChangeLog.md](/docs/ChangeLog.md)** add a description of the changes . See previous changes.
- on top of **[readme.md](/readme.md)**, change:
  - version number and
  - date
- In **Source\Changelog_Thor.txt**, copy the changes of the recent (and only the recent) version. File is available in the project.
- If there is more documentation, alter appropriate files
- In the footer of any file changed, alter date.

- In **Source\BuildGoFish.prg**, available in the project, change at top the 
  - `lcVersion` and 
  - `lcBuild` variables   

- run **Source\BuildGoFish.prg** to
  - update several files
  - build GoFish.app,
   - and generate the text equivalents for all VFP binary files (SCX, VCX, DBF, etc.) using FoxBin2PRG.
   > **Note: use VFP 9 SP2 rather than VFP Advanced to run BuildGoFish.prg since GoFish.app must be built in VFP rather than VFPA for compatibility for all users.**

- **close the project**
- In File Explorer, right-click **Source\BuildCloudZip.ps1** and choose *Run with PowerShell* to create Source.zip.
- Commit the changes.
- Push to your fork.
- Create a pull request; ensure the description clearly describes the problem and solution or the enhancement.

----
Last changed: 2023-04-15