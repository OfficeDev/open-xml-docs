# Adding samples

To add a sample, run the the following:

```powershell
./add-sample.ps1 [word|presentation|spreadsheet] sample-name
```

This will scaffold out the projects in a common layout. Samples should be a single file, with the entrypoint being the command line so that `dotnet run -- [args]` can be used in the docs.

In the future, we expect to set up .editorconfig/stylecop to enforce a shared style across the samples, but for now, the goal is to move the inline samples to this compilable solution.

General changes to move a sample:

- Many examples give details on how to open a project; this can be removed
- Sections about what `Dispose/Close/etc` is can be removed - this is an artifact from before `using` was common
- Samples currently have a "How the Sample Works" section followed by the actual sample. Going forward, this will be collapsed to just the sample - any comments required will be in the cs/vb
- Many users use the VB examples, so we will maintain them. Using docfx tabs allows us to hide the languages not needed by a viewer