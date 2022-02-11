import * as React from 'react';
import styles from './Attachments.module.scss';
import { IAttachmentsProps } from './IAttachmentsProps';

import { ListItemAttachments } from '@pnp/spfx-controls-react/lib/ListItemAttachments';
import { getControlledDerivedProps } from 'office-ui-fabric-react/lib/Foundation';
import {SPHttpClient} from "@microsoft/sp-http";


export default function Attachments (props: IAttachmentsProps){

  React.useEffect(()=>{
    console.log(props.context.pageContext.listItem.id);

    const responseUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/Lists/GetByTitle('Site Pages')/items(33)'`;
    const response = props.context.spHttpClient.get(responseUrl, SPHttpClient.configurations.v1).then(r => r.json());
    response.then(result => console.log(result['Description']));

  }, []);


  return (
    <div className={ styles.attachments }>
      Test Attachments...
      {/* <ListItemAttachments listId='d021c002-c529-4a37-8a06-e0e16276cd2f'
        context={this.props.context}
        disabled={false} 
      /> */}
    </div>
  );
}


