using System;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Implementations;
using Skrypton.Tests.RuntimeSupport.Implementations;

namespace Skrypton.Tests.RuntimeSupport.Components.OutlookApplication;

[SourceClassName("Application")] // "Outlook.Application"
[ComVisible(true)]
public sealed class MyOutlookApplicationClass : IReflectOnClrType
{
    // "Interop.Microsoft.Office.Interop.Outlook.dll"  <Guid>00062fff-0000-0000-c000-000000000046</Guid>
    // In the Outlook COM type library (OUTL.TLB)
    // use VisualStudio [command prompt] or
    // c:\Program Files\Microsoft Visual Studio\18\Enterprise\Common7\Tools>OleView.exe
    // => Go to 'Type Libraries', Scroll down to 'Microsoft Outlook x.x Object Library' - DoubleClick
    // As Administrator : "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\oleview.exe"
    private readonly IRuntimeHost _runtimeHost;

    public MyOutlookApplicationClass(IRuntimeHost runtimeHost)
    {
        _runtimeHost = runtimeHost ?? throw new ArgumentNullException(nameof(runtimeHost));
    }

    //[DispId(266)]
    //public object CreateItem()
    //{
    //    return CreateItem(OlItemType.olMailItem);
    //}

    [DispId(266)]
    public object CreateItem([In] object item)
    {
        OlItemType ItemType = (OlItemType)Enum.ToObject(typeof(OlItemType), item);
        //OlItemType ItemType = (OlItemType)item;
        if (ItemType == OlItemType.olContactItem)
        {
            return new MyOutlookContactItemClass();
        }
        if (ItemType == OlItemType.olMailItem)
        {
            return new MyOutlookMailItemClass();
        }
        throw new NotImplementedException($"ItemType:{ItemType}");
    }

    [DispId(272)]
    public _NameSpace GetNamespace([In] string Type) // Type:MAPI
    {
        if (Type == "MAPI")
        {
            return new MyOutlookMAPISession();
        }
        throw new NotImplementedException($"Type:{Type}");
    }
}

internal sealed class MyOutlookMailItemClass : IReflectOnClrType
{
    internal MyOutlookMailItemClass()
    {
    }
    [DispId(55)]
    public string Subject { get; set; }

    [DispId(3587)]
    public string CC { get; set; }
    [DispId(3588)]
    public string To { get; set; }


    [DispId(61606)]
    public void Display()
    {

    }
}

public enum OlItemType
{
    olMailItem = 0,
    olAppointmentItem = 1,
    olContactItem = 2,
    olTaskItem = 3,
    olJournalItem = 4,
    olNoteItem = 5,
    olPostItem = 6,
    olDistributionListItem = 7,
    olMobileItemSMS = 11,
    olMobileItemMMS = 12
}

//[ComVisible(true)]
internal sealed class MyOutlookContactItemClass : IReflectOnClrType
{
    public MyOutlookContactItemClass()
    {

    }
    [DispId(14854)]
    public string FirstName
    {
        get;
        set;
    }
    [DispId(14865)]
    public string LastName
    {
        get;
        set;
    }
    [DispId(14870)] public string CompanyName { get; set; }
    [DispId(14871)] public string JobTitle { get; set; }
    [DispId(32837)] public string BusinessAddressStreet { get; set; }
    [DispId(32838)] public string BusinessAddressCity { get; set; }
    [DispId(32839)] public string BusinessAddressState { get; set; }
    [DispId(32841)] public string BusinessAddressCountry { get; set; }
    [DispId(32840)] public string BusinessAddressPostalCode { get; set; }
    [DispId(14856)] public string BusinessTelephoneNumber { get; set; }
    [DispId(14884)] public string BusinessFaxNumber { get; set; }
    [DispId(32899)] public string Email1Address { get; set; }
    [DispId(14876)] public string MobileTelephoneNumber { get; set; }
    [DispId(61512)]
    public void Save()
    {

    }

    [DispId(61606)]
    public void Display()//[Optional][In] object Modal)
    {
        //throw new NotImplementedException($"Modal:{Modal}");
    }
}

public enum OlDefaultFolders
{
    olFolderDeletedItems = 3,
    olFolderOutbox = 4,
    olFolderSentMail = 5,
    olFolderInbox = 6,
    olFolderCalendar = 9,
    olFolderContacts = 10,
    olFolderJournal = 11,
    olFolderNotes = 12,
    olFolderTasks = 13,
    olFolderDrafts = 16,
    olPublicFoldersAllPublicFolders = 18,
    olFolderConflicts = 19,
    olFolderSyncIssues = 20,
    olFolderLocalFailures = 21,
    olFolderServerFailures = 22,
    olFolderJunk = 23,
    olFolderRssFeeds = 25,
    olFolderToDo = 28,
    olFolderManagedEmail = 29,
    olFolderSuggestedContacts = 30
}
public interface _NameSpace
{
}
internal sealed class MyOutlookMAPISession : IReflectOnClrType, _NameSpace
{
    [DispId(8459)]
    public MAPIFolder GetDefaultFolder([In] OlDefaultFolders FolderType)
    {
        if (FolderType == OlDefaultFolders.olFolderContacts)
        {
            return new MyOutlookMAPIFolderContacts();
        }
        throw new NotImplementedException($"FolderType:{FolderType}");
    }
}

internal interface MAPIFolder
{

}
interface _Items
{
}
internal sealed class MyOutlookMAPIFolderContacts : IReflectOnClrType, MAPIFolder
{
    private readonly ContactsItems _items = new ContactsItems();
    public MyOutlookMAPIFolderContacts()
    {
    }

    [DispId(12544)]
    public _Items Items
    {
        get
        {
            return _items;
        }
    }

    private sealed class ContactsItems : IReflectOnClrType, _Items
    {
        public ContactsItems()
        {
        }

        [DispId(98)]
        public object Find([In] string Filter) // Filter:[lastname] = ''  And [firstname] = ''
        {
            //throw new NotImplementedException($"Filter:{Filter}");
            return null;
        }
    }
}