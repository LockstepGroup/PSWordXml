# Contributing

If you use [VS Code](https://code.visualstudio.com/), this repo containes a settings file to apply certain formatting to match the desired guidelines. As a general rule of thumb this [Practice and Style Guide](https://github.com/PoshCode/PowerShellPracticeAndStyle) should be used.

## Updating Module Documentation
Module documentation is generated/maintained with [platyPS](https://github.com/PowerShell/platyPS). The following is an example assuming you run the commands from the repo directory.

```powershell
# install and import platyPS
Install-Module -Name platyPS -Scope CurrentUser
Import-Module platyPS

# import module for documentation
Import-Module PSWordXml

# update/generate markdown files for cmdlet help
Update-MarkdownHelpModule -Path ./docs/cmdlets
```

Edit the markdown files generated in `./docs/cmdlets` to contain the help you want.

```powershell
# convert markdown help to xml for packaging with module
New-ExternalHelp -Path ./docs/cmdlets/ -OutputPath ./PSWordXml/en-US/ -Force
```

## Build/Deploy module

Once your changes have been approved by a peer and merged into master, you're ready to build the module and deploy to PSGallery. Even though you peer most definitley made sure that all tests were passing before merging, you should do so as well just in case.

```powershell
# from the repo directory
Invoke-Pester
```

If anything fails, stop and go back to the drawing board. Otherwise, carry on.

```powershell
# Build the module, this will run through the tests, update the module manifest version, and some other cleanup
. ./build.ps1
```

Now you're ready to push these changes and trigger appveyor to deploy to PSGallery. Commit the changes made from the build process and push. The commit message needs to contain the string '!deploy', this will trigger appveyor's process.

```bash
# commit with !deploy in commit message
git commit -am '[descriptive commit message] !deploy'

# push to github
git push
```

## Git primer

Most of the commands you need to use git for this repo.  Feel free to use the github desktop client, but don't ask me how.

### Generic workflow for adding code

```bash
# initial download of the repo, get link from the green "Clone or download" button on github
git clone [url]

# create a branch (and switch context to it) to make changes on, never work on master
git checkout [branch-name]

# at this point you can make changes and save files without worrying about master branch.

# add all new files to your branch, if you make new files and don't run this, you will not track these files with git
git add .

# commit changes to your branch, you should run this often, not after making tons and tons of changes.
# ie:
# - added a new switch to a cmdlet
# - updated some doc
# - removed a typo
# - added a test
git commit -am "[descriptive message for your commit]"

# push changes back to github, use this to make code available to other users for testing/review/etc
git push origin [branch-name]

# pull changes from github to update your local branch, for instance, if someone makes changes to code you just pushed
git pull
```

### Submit pull request to merge new code

If you're ready to merge your changes into master (production), then you need to create a pull request. Do this with the following process.

1. Login to github and navigate to the desired repository.
2. Click the "branches" button for your repo.
3. Find your branch under "Active branches" and click "New pull request".
4. Write up a description of what your new code does and click "Create pull request"

This will alert other contributers that you are ready for code review. After completing a code review to satisfcation, the code will be merged to master. At this point, if you have more changes for the current branch, then repeat the process. If you are done with the branch, then you have some cleanup to do.

### Delete branch from local repo

```bash
# switch to master before deleting the branch
git checkout master

# delete your local copy of the branch
git branch -D [branch-name]

# update your copy of master from github
git pull
```

### Delete branch from github

1. Login to github and navigate to the desired repository.
2. Click the "branches" button for your repo.
3. Find your branch under "Active branches" and click the delete icon.

### Helpful links
[posh-git](https://github.com/dahlbyk/posh-git): adds some nice formatting to your powershell prompt for git status
[cheatsheet](https://github.github.com/training-kit/downloads/github-git-cheat-sheet/).