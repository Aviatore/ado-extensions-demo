import * as React from 'react';
import * as SDK from 'azure-devops-extension-sdk';

import { IWorkItemFormService, WorkItemTrackingServiceIds } from 'azure-devops-extension-api/WorkItemTracking';
import { IExtensionDataService, IExtensionDataManager, CommonServiceIds } from 'azure-devops-extension-api';

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

    const [checked, setChecked] = React.useState(false);
    const personal = React.useRef(false);
    const userId = React.useRef<string | null>(null);

    console.log("render");

    const onCheckHandler = (e: React.ChangeEvent<HTMLInputElement>) => {
        personal.current = !personal.current;
        setChecked(personal.current);
        loadData();
    }

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

        const witId = workItemId.current;
        const id = personal.current ? `${userId.current}.${workItemId.current.toString()}` : workItemId.current.toString();

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

    const onChangeHandler = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
        setStorageInputValue(e.target.value)
    }
    
    const onSave = (e: React.MouseEvent<HTMLInputElement>) => {
        saveData(storageInputValue);
    }

    const loadData = async () => {
        const dataManager = await getDataManager();
        const id = personal.current ? `${userId.current}.${workItemId.current.toString()}` : workItemId.current.toString();

        try
        {
            const document: IData = await dataManager.getDocument("myExt", id);

            setStorageInputValue(document.data);
        } catch {
            console.log(`Cannot get a document. Collection 'myExt', Work Item ID: '${id}'`);
            setStorageInputValue("");
        }
    }
    
    const registerEvents = () => {
        SDK.register(SDK.getContributionId(), () => {
            return {
                onLoaded: async() => {                    
                    maxWordCount.current = SDK.getConfiguration().witInputs["MaxWordCount"]; 
                    
                    const user = SDK.getUser();
                    userId.current = user.id;
                    console.log(user);

                    fieldName.current = SDK.getConfiguration().witInputs["FieldName"];

                    const wifs = await getWorkItemFormService();
                    workItemId.current = await wifs.getId();
                    
                    getWordCount();

                    loadData();
                },
                onFieldChanged: () => {
                    getWordCount();
                },
                onSaved: () => {},
                onReset: () => {},
                onRefreshed: () => {}
            }
        });
    }

    React.useEffect(() => {
        SDK.init().then(() => {
            registerEvents();
        });
    }, []);
   

    return (
        <>
            <div>
                <span>Max word count: {wordStat.maxWordCount}</span>
                
            </div>
            <div>
                <span>Current word count: {wordStat.wordCount}</span>
            </div>
            <div>
                <label htmlFor='storage'>My Notes:</label><br />
                <textarea id='storage' onChange={onChangeHandler} value={storageInputValue} /><br />
                <label htmlFor='personal'>Personal</label>
                <input id='personal' type='checkbox' checked={personal.current} onChange={onCheckHandler} /><br />
                <input type='button' onClick={onSave} value='Save' />                
            </div>
        </>
    );
}

export default App;