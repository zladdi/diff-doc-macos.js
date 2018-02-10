#!/usr/bin/env osascript -l JavaScript

//  diff-doc-macos.js
//
//  A partial port of TortoiseSVN/TortoiseGit diff-doc.js to Open Scripting Architecture (OSA) on macOS.
//  For the source script, see https://github.com/TortoiseGit/TortoiseGit/blob/master/contrib/diff-scripts/diff-doc.js
//
//  The OpenOffice portion was not ported.
//
//  This file is distributed under the GNU General Public License.
//
//  Author: Zlatko Franjcic

function run(argv)
{   
    ObjC.import('stdlib')
    ObjC.import('stdio')

    var word
    var sTempDoc, sBaseDoc, sNewDoc
    var destination

    // Microsoft Office versions for Microsoft Windows OS
    const vOffice2000 = 9, vOffice2002 = 10, //vOffice2003 = 11,
        vOffice2007 = 12, vOffice2013 = 15

    // WdCompareTarget
    //const wdCompareTargetSelected = 'compare target selected'
    //const wdCompareTargetCurrent = 'compare target current'
    const wdCompareTargetNew = 'compare target new'
    // WdViewType
    const wdMasterView = 'master view'
    const wdNormalView = 'normal view'
    const wdOutlineView = 'outline view'
    const wdReadingView = 'WordNote view' // 7

    // WdSaveOptions
    const wdDoNotSaveChanges = 'no'
    //const wdPromptToSaveChanges = 'ask'
    //const wdSaveChanges = 'yes'

    // WdOpenFormat
    const wdOpenFormatOpenFormatAuto = 'open format auto'

    argc = argv.length 
    if (argv.length < 2)
    {
        var scriptApp = Application.currentApplication()
        scriptApp.includeStandardAdditions = true
        const basename = $.NSString.alloc.initWithUTF8String(
                            scriptApp.pathTo(this))
                            .lastPathComponent.UTF8String
        $.printf('Usage: %s <absolute-path-to-base.doc> <absolute-path-to-new.doc>\n', basename)
        $.exit(1)
    }

    sBaseDoc = argv[0]
    sNewDoc = argv[1]

    if (!$.NSFileManager.defaultManager.fileExistsAtPath(sBaseDoc))
    {
        $.printf('File %s does not exist. Cannot compare the documents.\n', sBaseDoc)
        $.exit(1)
    }
    
    if (!$.NSFileManager.defaultManager.fileExistsAtPath(sNewDoc))
    {
        $.printf('File %s does not exist. Cannot compare the documents.\n', sNewDoc)
        return 1
    }
    try {
        word = Application('com.microsoft.Word')
        if (parseInt(word.version()) >= vOffice2013)
        {
            if (!$.NSFileManager.defaultManager.isWritableFileAtPath(sBaseDoc))
            {
                // reset read-only attribute
                $.NSFileManager.defaultManager.setAttributesOfItemAtPathError($({NSFileImmutable: $.NSNumber.numberWithBool(false)}), sBaseDoc, $())
            }
        }
    }
    catch(ex)
    {
        $.printf('You must have Microsoft Word installed to perform this operation.\n')
        $.exit(1)
    }
    
    if (parseInt(word.version()) >= vOffice2007)
    {
        sTempDoc = sNewDoc
        sNewDoc = sBaseDoc
        sBaseDoc = sTempDoc
    }
    
    // The 'visible' property does not exist in this interface
    //word.visible
    
    // Open the new document
    try
    {
        destination = word.open(null, {
            fileName:sNewDoc,
            confirmConversions:true,
            readOnly:(parseInt(word.version()) < vOffice2013),
            addToRecentFiles:false,
            repair:false,
            showingRepairs:false,
            passwordDocument:null,
            passwordTemplate:null,
            revert:false,
            writePassword:null,
            writePasswordTemplate:null,
            fileConverter:wdOpenFormatOpenFormatAuto
            })
    }
    catch(ex)
    {
        try
        {
            // open empty document to prevent bug where first Open() call fails
            destination = word.activeDocument
            destination = word.open(null, {
                fileName:sNewDoc,
                confirmConversions:true,
                readOnly:(parseInt(word.version()) < vOffice2013),
                addToRecentFiles:false,
                repair:false,
                showingRepairs:false,
                passwordDocument:null,
                passwordTemplate:null,
                revert:false,
                writePassword:null,
                writePasswordTemplate:null,
                fileConverter:wdOpenFormatOpenFormatAuto
                })
        }
        catch(ex)
        {
            $.printf('Error opening %s\n', sNewDoc)
            // Quit
            $.exit(1)
        }
    }

            
    // If the Type property returns either wdOutlineView or wdMasterView and the Count property returns zero, the current document is an outline.
    
    if (((destination.activeWindow.view.viewType == wdOutlineView) || ((destination.activeWindow.view.viewType == wdMasterView) || (destination.activeWindow.view.viewType == wdReadingView))) && (destination.subdocuments.count == 0))
    {
        // Change the Type property of the current document to normal
        destination.activeWindow.view.setViewType(wdNormalView)
    }
    
    // Compare to the base document
    if (parseInt(word.version()) <= vOffice2000)
    {
        // Compare for Office 2000 and earlier
        try
        {
            // Contrary to the original TortoiseSVN/Git script, we cannot use duck typing -> comment out this line,
            // as we only support the newer interface below
            //[destination comparePath:sBaseDoc]
            $.printf('Warning: Office versions up to Office 2000 are not officially supported.\n')
            destination.comparePath(sBaseDoc, {
                authorName: 'Comparison',
                target: wdCompareTargetNew,
                detectFormatChanges: true,
                ignoreAllComparisonWarnings: true,
                addToRecentFiles: false
            })
            
        }
        catch(ex)
        {
            $.printf('Error comparing %s and %s\n', sBaseDoc, sNewDoc)
            // Quit
            $.exit(1)
        }
    }
    else
    {
        // Compare for Office XP (2002) and later
        try
        {
            destination.comparePath(sBaseDoc, { 
                authorName: 'Comparison',
                target: wdCompareTargetNew,
                detectFormatChanges: true,
                ignoreAllComparisonWarnings: true,
                addToRecentFiles: false
            })
        }
        catch(ex)
        {
            $.printf('Error comparing %s and %s\n', sBaseDoc, sNewDoc)
            // Close the first document and quit
            destination.closeSaving(wdDoNotSaveChanges, {savingIn:null})
            $.exit(1)
        }
    }
    
    // Show the comparison result
    if (parseInt(word.version()) < vOffice2007)
    {
        word.activeDocument.windows[0].setVisible(true)
    }
    
    // Mark the comparison document as saved to prevent the annoying
    // 'Save as' dialog from appearing.
    word.activeDocument.setSaved(true)
    
    // Close the first document
    if (parseInt(word.version()) >= vOffice2002)
    {
        destination.closeSaving(wdDoNotSaveChanges, {savingIn: null})
    }

    $.exit(0)
}
    