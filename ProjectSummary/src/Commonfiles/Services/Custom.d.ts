 declare module "CustomJS"{
    interface ICustomjs{
        load():void;
        fnclear():void;
       
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