 declare module "CustomJS"{
    interface ICustomjs{
        load():void;
        fnAlert(text:string):void;
        fnUploadAttachments(siteURL:string,itemID:string):void;
    }

    var CustomJS :ICustomjs; 
    export = CustomJS;
}


// declare module "MultiFile" {
//     interface IMultiFile{
//         MultiFile():void;
//     }

//     var MultiFile :IMultiFile; 
//     export = MultiFile;
// }