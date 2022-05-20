use std::{
    ffi::c_void,
    mem::MaybeUninit,
};

use anyhow::Result;
use windows::{
    core::GUID,
    Win32::{
        Foundation::HANDLE,
        System::Com,
        System::Com::{CoInitializeEx},
        UI::Shell::{
            FOLDERID_RecycleBinFolder, IShellItem,
            SHGetKnownFolderItem, KF_FLAG_DEFAULT, SIGDN_NORMALDISPLAY, IEnumShellItems, BHID_EnumItems,
        },
    },
};

// Trying to convert this old Raymond Chen article to Rust: https://devblogs.microsoft.com/oldnewthing/20110831-00/?p=9763
// attempting to list items in the Recycle Bin
fn main() -> Result<()> {
    unsafe {
        CoInitializeEx(std::ptr::null(), Com::COINIT_APARTMENTTHREADED)?;

        // Couldn't find this in the `windows` crate
        // {43826D1E-E718-42EE-BC55-A1E261C37BFE}
        let IID_IShellItem = GUID::from_values(
            0x43826d1e,
            0xe718,
            0x42ee,
            [0xbc, 0x55, 0xa1, 0xe2, 0x61, 0xc3, 0x7b, 0xfe],
        );

        // TODO: make sure we free this
        let mut recycle_bin = MaybeUninit::<*mut IShellItem>::uninit();

        SHGetKnownFolderItem(
            &FOLDERID_RecycleBinFolder,
            KF_FLAG_DEFAULT,
            HANDLE::default(),
            &IID_IShellItem,
            recycle_bin.as_mut_ptr() as *mut *mut c_void,
        )?;

        let recycle_bin = recycle_bin.assume_init();
        
        // let name = recycle_bin.as_ref().unwrap().GetDisplayName(SIGDN_NORMALDISPLAY)?;

        // this crashes with a STATUS_ACCESS_VIOLATION
        let pesi: IEnumShellItems = (*recycle_bin).BindToHandler(None, &BHID_EnumItems)?;

        dbg!(pesi);
    }

    println!("Done!");
    Ok(())
}
