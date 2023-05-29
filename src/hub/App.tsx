import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';

import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { IExtensionDataService, IExtensionDataManager, CommonServiceIds } from 'azure-devops-extension-api';

import { ObservableValue, IReadonlyObservableValue } from "azure-devops-ui/Core/Observable";
import { Observer } from "azure-devops-ui/Observer";
import { Tab, TabBar, TabSize } from "azure-devops-ui/Tabs";
import { TextField, TextFieldWidth } from "azure-devops-ui/TextField";
import { Button } from "azure-devops-ui/Button";
import { Checkbox } from "azure-devops-ui/Checkbox";

interface IWordStats {
    maxWordCount: number,
    wordCount: number
}

interface IData {
    id: string,
    data: string
}

const App = () => {
    const [wordStat, setWordStat] = React.useState<IWordStats>({
        maxWordCount: 0,
        wordCount: 0
    });

    const maxWordCount = React.useRef(0);
    const fieldName = React.useRef("");
    const wordCount = React.useRef(0);
    const workItemId = React.useRef(0);
    const [storageInputValue, setStorageInputValue] = React.useState("");
    const selectedTabId = React.useRef(new ObservableValue("tab1"));
    const checkbox = React.useRef(new ObservableValue<boolean>(false));
    const userId = React.useRef<string | null>(null);

    console.log("render");

    const getWorkItemFormService = async() => {
        return await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
    }

    const getWordCount = async () => {
        const wifs = await getWorkItemFormService();

        let wc = 0;
        const fieldContent = await wifs?.getFieldValue(fieldName.current.toString(), {returnOriginalValue: false});
        const textRaw = fieldContent?.toString().replace(/<[^>]+>/g, '')
        const words = textRaw?.split(/\s+/).filter(p => p);
        
        if (words) {
            wc = words.length;
        }
        
        wordCount.current = wc;
        checkWordCount();
        setWordStat({
            maxWordCount: parseInt(maxWordCount.current.toString()),
            wordCount: wc
        });
    }

    const checkWordCount = async () => {
        const wifs = await getWorkItemFormService();

        if (wordCount.current > maxWordCount.current) {
            await wifs?.setError(`Description contains ${wordCount.current} words but should have less than ${maxWordCount.current} words`);
        } else {
            await wifs?.clearError();
        }
    }

    const getDataManager = async () => {
        const extensionDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const id = SDK.getExtensionContext().id;
        const accessToken = await SDK.getAccessToken();
        const extensionDataManager = extensionDataService.getExtensionDataManager(id, accessToken);

        return extensionDataManager;
    }

    const saveData = async (newData: string) => {
        const dataManager = await getDataManager();

        let document: IData;
        const id = checkbox.current.value ? `${userId.current}.${workItemId.current.toString()}` : workItemId.current.toString();

        try
        {
            document = await dataManager.getDocument("myExt", id);
            document.data = newData;

            console.log("Editing existing storage ...");
            await dataManager.updateDocument("myExt", document);
        } catch {
            console.log("Creating a new storage ...");
            const newDocument: IData = {
                id: id,
                data: newData
            }
            dataManager.createDocument("myExt", newDocument);
        }

        setStorageInputValue(newData);
    }

    const onChangeHandler = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>, value: string) => {
        setStorageInputValue(e.target.value)
    }
    
    const onSave = () => {
        saveData(storageInputValue);
    }

    const loadData = async () => {
        const dataManager = await getDataManager();
        const id = checkbox.current.value ? `${userId.current}.${workItemId.current.toString()}` : workItemId.current.toString();

        try
        {
            const document: IData = await dataManager.getDocument("myExt", id);

            setStorageInputValue(document.data);
        } catch {
            console.log(`Cannot get a document. Collection 'myExt', Work Item ID: '${id}'`);
            setStorageInputValue("");
        }
    }

    const onSelectedTabChanged = (newTabId: string) => {
        selectedTabId.current.value = newTabId;
    }
    
    const registerEvents = () => {
        SDK.register(SDK.getContributionId(), () => {
            return {
                onLoaded: async() => {                    
                    maxWordCount.current = SDK.getConfiguration().witInputs["MaxWordCount"];                    

                    fieldName.current = SDK.getConfiguration().witInputs["FieldName"];

                    const wifs = await getWorkItemFormService();
                    workItemId.current = await wifs.getId();

                    const user = SDK.getUser();
                    userId.current = user.id;
                    console.log(user);
                    
                    getWordCount();

                    loadData();
                },
                onFieldChanged: () => {
                    getWordCount();
                }
            }
        });
    }

    React.useEffect(() => {
        SDK.init().then(() => {
            registerEvents();
        });
    }, []);

    const wordCountElement: () => JSX.Element = () => {
        return (
            <>
                <div>
                    <span>Max word count: {wordStat.maxWordCount}</span>
                    
                </div>
                <div>
                    <span>Current word count: {wordStat.wordCount}</span>
                </div>
            </>
        )
    }

    const onCheckHandler = (e: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement> | undefined, checked: boolean) => {
        checkbox.current.value = checked;
        loadData();
    }

    const notesElement: () => JSX.Element = () => {
        return (
            <div>
                <TextField 
                    ariaLabel='Aria label'
                    onChange={onChangeHandler}
                    value={storageInputValue} 
                    multiline
                    rows={3}
                    width={TextFieldWidth.standard}/>
                <Checkbox
                    onChange={onCheckHandler}
                    checked={checkbox.current}
                    label="Personal"
                /><br />
                <Button text='Save' primary={true} onClick={onSave} />
            </div>
        )
    }
   

    return (
        <>
            <TabBar onSelectedTabChanged={onSelectedTabChanged}
                    selectedTabId={selectedTabId.current}
                    tabSize={TabSize.Tall}>
                <Tab name='Word Count' id='tab1' />
                <Tab name='Notes' id='tab2' />
            </TabBar>
            <Observer selectedTabId={selectedTabId.current}>
                {
                    (props: {selectedTabId: string}) => {
                        switch (props.selectedTabId) {
                            case "tab1":
                                return wordCountElement();
                                break;
                            case "tab2":
                                return notesElement();
                                break;
                        }
                    }
                }
            </Observer>
        </>
    );
}

export default App;