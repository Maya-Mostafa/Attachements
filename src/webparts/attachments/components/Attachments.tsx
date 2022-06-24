import * as React from 'react';
import styles from './Attachments.module.scss';
import { IAttachmentsProps } from './IAttachmentsProps';
import {getPgAttachments, getFileIcon, getOpenInBrowserLink, deleteAttachment, isUserManage} from '../../Services/Requests';
import {Icon, Dialog, DialogFooter, PrimaryButton, DefaultButton, Spinner, SpinnerSize, Overlay, DialogType} from 'office-ui-fabric-react';
import { useBoolean } from '@uifabric/react-hooks';

export default function Attachments (props: IAttachmentsProps){

  const [attachments, setAttachments] = React.useState([]);
  const [attachItem, setAttachItem] = React.useState(null);
  const [isDataLoading, { toggle: toggleIsDataLoading }] = useBoolean(false);
  const [hideDeleteDialog, {toggle: toggleHideDeleteDialog}] = useBoolean(true);

  const deleteDialogContentProps = {
    type: DialogType.close,
    title: "Delete File"
  }; 

  React.useEffect(()=>{
    getPgAttachments(props.context).then(r=> {console.log(r.value);setAttachments(r.value);});
  }, []);

  const onDeleteClickHandler = (item: any) => {
    setAttachItem(item);
    toggleHideDeleteDialog();
  };
  const onDeleteAttachment = () => {
    toggleHideDeleteDialog();
    toggleIsDataLoading();
    deleteAttachment(props.context, attachItem).then(r1 => {
      getPgAttachments(props.context).then(r=> {
        setAttachments(r.value);
        toggleIsDataLoading();
      });      
    });
  };

  return (
    <div className={ styles.attachments }>
      <ul className={styles.attachmentsList}>
        {attachments.map(item => {
          return (
            <li>
              <a className={styles.attachmentLinkOpen} href={getOpenInBrowserLink(props.context, item)} target="_blank" data-interception="off">
                <i>
                  <img width="20px" src={getFileIcon(item.ServerRelativeUrl)} alt={item.Name}/>
                </i>
                <span>{item.Name.substring(0,item.Name.lastIndexOf('.'))}</span>
              </a>
              <a className={styles.attachmentLinkDownload} href={`${item.ServerRelativeUrl}`} title='Download' download>
                <Icon iconName='Download' />
              </a>
              {isUserManage &&
                <a className={styles.attachmentDelete} onClick={() => onDeleteClickHandler(item)} title='Delete'>
                  <Icon iconName='Delete' />
                </a>
              }
            </li>
          );
        })}
      </ul>
      <Dialog
            hidden={hideDeleteDialog}
            onDismiss={toggleHideDeleteDialog} isBlocking={true}
            dialogContentProps={deleteDialogContentProps}>
            <p>Are you sure you want to delete this file? </p>
            {isDataLoading &&
              <>
                  <Overlay></Overlay>
                  <div>
                      <Spinner size={SpinnerSize.medium} label="Please Wait, Updating Attachments..." ariaLive="assertive" labelPosition="right" />
                  </div>
              </>
            }
            <DialogFooter>
                <PrimaryButton onClick={onDeleteAttachment} text="Yes" />
                <DefaultButton onClick={toggleHideDeleteDialog} text="No" />
            </DialogFooter>
        </Dialog>
        
    </div>
  );
}


