# Adding samples

## New Samples

To add a sample, run the the following:

```powershell
./add-sample.ps1 area name
```

This will create an initial scaffold for a sample and add it to the solution file.

## Migrate old samples

```powershell
./migrate-sample.ps1 path-to-md-file
```

This will do an inital extraction and clean up of the file, as well as add the code to the solution. Additional clean up will be necessary, but should be minimal.

General changes to move a sample:

- Many examples give details on how to open a project; this can be removed
- Sections about what `Dispose/Close/etc` is can be removed - this is an artifact from before `using` was common
- Samples currently have a "How the Sample Works" section followed by the actual sample. Going forward, this will be collapsed to just the sample - any comments required will be in the cs/vb
- Many users use the VB examples, so we will maintain them. Using docfx tabs allows us to hide the languages not needed by a viewer

## Code set up

In the future, we expect to set up .editorconfig/stylecop to enforce a shared style across the samples, but for now, the goal is to move the inline samples to this compilable solution.
