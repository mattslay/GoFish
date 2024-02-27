# How to contribute to GoFish

## Report a bug
- Please check [issues](https://github.com/VFPXGoFish/issues) if the bug has already been reported.
- If you're unable to find an open issue addressing the problem, open a new one. Be sure to include a title and clear description, as much relevant information as possible, and a code sample or an executable test case demonstrating the expected behavior that is not occurring.

## Fix a bug or add an enhancement
- Fork the project: see this [guide](https://www.dataschool.io/how-to-contribute-on-github/) for setting up and using a fork.
  - If already forked, pull the recent state, or get most recent version otherwise.
- Make whatever changes are necessary.
- In **[docs\ChangeLog.md](/docs/ChangeLog.md)** add a description of the changes . See previous changes.

*** 
The steps in this section only apply if you intend to merge your changes immediately into the master repository for Thor. This can only happen if you have access to do so and you are sure that your changes do not need any further testing by others. If you do not intend to merge into the master repository or are in any way unsure what this means, skip all steps in this section.

1. In **Source\BuildGoFish.prg**, available in the project, update two files:
    * VersionNumber.txt
    * BuildNumber.txt
1. On top of **[readme.md](/readme.md)**, change:
    * Version number
    * Date
1. In **Source\Changelog_Thor.txt**, copy the changes of the recent (and only the recent) version. File is available in the project.
    * If there is more documentation, alter appropriate files
    * In the footer of any file changed, alter date.
1. Run **Source\BuildGoFish.prg** to
    * update several files
    * build GoFish.app
   * generate the text equivalents for all VFP binary files (SCX, VCX, DBF, etc.) using FoxBin2PRG.
   > **Note: use VFP 9 SP2 rather than VFP Advanced to run BuildGoFish.prg since GoFish.app must be built in VFP rather than VFPA for compatibility for all users.**
1.  **Close the project**
1. In File Explorer, right-click **Source\BuildCloudZip.ps1** and choose *Run with PowerShell* to create Source.zip.

*** 

- Commit the changes.
- Push to your fork.
- Create a pull request; ensure the description clearly describes the problem and solution or the enhancement.

----
Last changed: 2024-02-25 ![](../docs/pictures/vfpxpoweredby_alternative.gif)