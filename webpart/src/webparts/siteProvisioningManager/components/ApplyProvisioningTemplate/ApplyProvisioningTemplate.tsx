import * as React from 'react';
import styles from "../App/App.module.scss";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import AppContext from "../App/AppContext";
import { MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import * as Strings from "SiteProvisioningManagerWebPartStrings";
import { FilePond } from 'react-filepond';
import 'filepond/dist/filepond.min.css';

const SiteProvisioningTemplate = () => {
    const ctx = React.useContext(AppContext);
    const [webUrl, setWebUrl] = React.useState(ctx.webUrl);

    const onUploadCompleted = (response) => {
        ctx.updateMessageBarSettings({
            message: Strings.SuccessMessage,
            type: MessageBarType.success,
            visible: true
        });
    };

    const onUploadFailed = (response) => {
        ctx.updateMessageBarSettings({
            message: response,
            type: MessageBarType.error,
            visible: true
        });
    };

    const fileRemoved = () => {
        ctx.updateMessageBarSettings({
            message: "",
            type: MessageBarType.info,
            visible: false
        });
    };
    
    return (
        <div className={styles.pivotContainer}>
            <div hidden={ctx.isLoading}>
                <TextField disabled={!ctx.isGlobalAdmin} label="Site URL" value={webUrl} onChanged={(value) => setWebUrl(value)} />
                <br />
                <FilePond disabled={!ctx.isSiteOwner} onremovefile={fileRemoved} server={
                    {
                        url: ctx.appService.applyProvisioningTemplateUrl,
                        process: {
                            method: 'POST',
                            ondata: (formData) => {
                                formData.append('WebUrl', webUrl);
                                return formData;
                            },
                            onload: onUploadCompleted,
                            onerror: onUploadFailed
                        }
                    }
                } />
            </div>
        </div>
    );
};

export default SiteProvisioningTemplate;