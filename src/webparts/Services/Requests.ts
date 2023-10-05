
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPPermission } from "@microsoft/sp-page-context";
import {SPHttpClient, ISPHttpClientOptions} from "@microsoft/sp-http";

export const getPgAttachments = async (context: WebPartContext) =>{

    // console.log("context", context);

    const pageUrl = document.documentURI;
    const pageTitle = encodeURIComponent(pageUrl.substring(pageUrl.lastIndexOf('/')+1, pageUrl.lastIndexOf('.aspx')));

    // console.log("pageTitle", pageTitle);

    const responseUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('siteassets/sitepages/${pageTitle}')/files`;
    const response = await context.spHttpClient.get(responseUrl, SPHttpClient.configurations.v1).then(r => r.json());
    return response;
};

export const getFileIcon = (fileUrl: String) =>{
    let fileIconUrl = '';
    if (fileUrl.indexOf('.docx') !== -1 || fileUrl.indexOf('.doc') !== -1){
        fileIconUrl =  "https://modernb.akamai.odsp.cdn.office.net/files/fabric-cdn-prod_20220127.003/assets/item-types/20/docx.svg";    
    }
    else if (fileUrl.indexOf('.pdf') !== -1){
        fileIconUrl = 'https://modernb.akamai.odsp.cdn.office.net/files/fabric-cdn-prod_20220127.003/assets/item-types/20/pdf.svg';
    }
    else if (fileUrl.indexOf('.xls') !== -1 || fileUrl.indexOf('.xlsx') !== -1){
        fileIconUrl = 'https://modernb.akamai.odsp.cdn.office.net/files/fabric-cdn-prod_20220127.003/assets/item-types/20/xlsx.svg';
    }
    else if (fileUrl.indexOf('.css') !== -1 || fileUrl.indexOf('.js') !== -1 || fileUrl.indexOf('.html') !== -1 || fileUrl.indexOf('.htm') !== -1){
        fileIconUrl = 'https://modernb.akamai.odsp.cdn.office.net/files/fabric-cdn-prod_20220127.003/assets/item-types/20/code.svg';
    }
    else if (fileUrl.indexOf('.ppt') !== -1 || fileUrl.indexOf('.pptx') !== -1){
        fileIconUrl = 'https://modernb.akamai.odsp.cdn.office.net/files/fabric-cdn-prod_20220127.003/assets/item-types/20/pptx.svg';
    }
    else if (fileUrl.indexOf('.vsd') !== -1 || fileUrl.indexOf('.vsdx') !== -1){
        fileIconUrl = 'https://modernb.akamai.odsp.cdn.office.net/files/fabric-cdn-prod_20220127.003/assets/item-types/20/vsdx.svg';
    }
    else if (fileUrl.indexOf('onrnote') !== -1){
        fileIconUrl = 'https://modernb.akamai.odsp.cdn.office.net/files/fabric-cdn-prod_20220127.003/assets/item-types/20/onetoc.svg';
    }
    else if (fileUrl.indexOf('.url') !== -1){
        fileIconUrl = 'https://modernb.akamai.odsp.cdn.office.net/files/fabric-cdn-prod_20220127.003/assets/item-types/20/link.svg';
    }
    else if (fileUrl.indexOf('.png') || fileUrl.indexOf('.jpg') || fileUrl.indexOf('.bmp') || fileUrl.indexOf('.gif')){
        fileIconUrl = 'https://res-1.cdn.office.net/files/fabric-cdn-prod_20220127.003/assets/item-types/20/photo.svg';
    }

    return fileIconUrl;
};


export const getOpenInBrowserLink = (context: WebPartContext, obj: any) => {
    const officeExts = ['.doc', '.docx', '.ppt', '.pptx', '.xls', '.xlsx', '.vsd', '.vsdx'];
    const fileExt = obj.ServerRelativeUrl.substring(obj.ServerRelativeUrl.indexOf('.'));
    let fileUrl = obj.ServerRelativeUrl;

    console.log("obj", obj);

    if (officeExts.indexOf(fileExt) !== -1){
        fileUrl = `${context.pageContext.web.absoluteUrl}/_layouts/15/Doc.aspx?sourcedoc=%7B${obj.UniqueId}%7D&file=${obj.Name}&action=default&mobileredirect=true`;
    }
    else {
        fileUrl = `${context.pageContext.web.absoluteUrl}/SiteAssets/Forms/AllItems.aspx?id=${encodeURIComponent(obj.ServerRelativeUrl)}&parent=${obj.ServerRelativeUrl.substring(0, obj.ServerRelativeUrl.indexOf(obj.Name)-1)}`;
    }

    return fileUrl;
};

export const deleteAttachment = async (context: WebPartContext, item: any) => {

    const restUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${item.ServerRelativeUrl}')/recycle`;

    let spOptions: ISPHttpClientOptions = {
        headers:{
            Accept: "application/json;odata=nometadata", 
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
            "IF-MATCH": "*",
            // "X-HTTP-Method": "DELETE"         
        },
    };

    const _data = await context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, spOptions);
    if (_data.ok){
        console.log('Attachment is deleted!');
    }
};

export const isUserManage = (context: WebPartContext) : boolean =>{
    const userPermissions = context.pageContext.web.permissions,
        permission = new SPPermission (userPermissions.value);
    
    return permission.hasPermission(SPPermission.manageWeb);
};