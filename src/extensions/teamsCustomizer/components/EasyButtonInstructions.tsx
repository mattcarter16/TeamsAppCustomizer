import * as React from 'react';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { MessageBar, MessageBarType } from '@fluentui/react';


export interface IEasyButtonInstructionsProps {
    context: ApplicationCustomizerContext;
}

const EasyButtonInstructions = (props: IEasyButtonInstructionsProps) => {

    console.log(props.context.pageContext.list.title);
    if (props.context.pageContext.list.title == 'Documents') {

        return (
            <div>
                <MessageBar>
                    <b>MS Teams Connected Storage</b>
                </MessageBar>

                <MessageBar
                    messageBarType={MessageBarType.warning}>
                    <b>Move to keynet instructions: </b>
                        <ol type="1">
                            <li>Select document(s) or folder(s) you want to move</li>
                            <li>Click on "Move to keynet" in the library action bar</li>
                            <li>Complete steps in dialog</li>
                        </ol>
                </MessageBar>
            </div>
        );
    } else {
        return (
            <div>
                <MessageBar>
                    <b>MS Teams Connected Site</b>
                </MessageBar>
            </div>
        );
    }

};

export default EasyButtonInstructions;
