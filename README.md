# Example Azure DevOps extension

The repo contains an exemplary extension for Azure DevOps. The extension (after installation in the ADO organization) can be added to any work item as a custom control. The extension provides two functionalities:
- Form validation - it counts words in the selected input field and triggers an error message when the word count exceeds a threshold value.
- Saving notes in the dedicated input field. By selecting the checkbox 'Personal' user can decide whether the saved data should not be visible to other users.

The repo contains two branches
- raw-version - contains all functionalities described above
- ms-ui-version - the upgraded version of the branch 'raw-version' that utilizes ui components provided by Microsoft ([azure-devops-ui](https://developer.microsoft.com/en-us/azure-devops/)).

Useful links
- [Starting project](https://github.com/microsoft/azure-devops-extension-hot-reload-and-debug)
- [VS Marketplace](https://marketplace.visualstudio.com/)
- [Extension samples](https://learn.microsoft.com/en-us/azure/devops/extend/develop/samples-overview?toc=%2Fazure%2Fdevops%2Fmarketplace-extensibility%2Ftoc.json&view=azure-devops)
- [Create a publisher](https://marketplace.visualstudio.com/manage/createpublisher?managePageRedirect=true)
- [Manage publishers](https://marketplace.visualstudio.com/manage)
- [azure-devops-extension-sdk (documentation)](https://learn.microsoft.com/en-us/javascript/api/azure-devops-extension-sdk/)
- [azure-devops-extension-api (documentation)](https://learn.microsoft.com/en-us/javascript/api/azure-devops-extension-api/)
- [azure-devops-ui (documentation)](https://developer.microsoft.com/en-us/azure-devops/)
- [IWorkItemFormService](https://learn.microsoft.com/en-us/javascript/api/azure-devops-extension-api/iworkitemformservice)