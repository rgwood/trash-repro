use std::{
    ffi::{c_void, OsString},
    mem::MaybeUninit, os::windows::prelude::OsStringExt, slice,
};

use anyhow::Result;
use windows::{
    core::{GUID, PWSTR, PCWSTR, Interface},
    Win32::{
        Foundation::HANDLE,
        System::Com::{self, CoTaskMemFree},
        System::Com::{CoInitializeEx},
        UI::Shell::{
            FOLDERID_RecycleBinFolder, IShellItem,
            SHGetKnownFolderItem, KF_FLAG_DEFAULT, SIGDN_NORMALDISPLAY, IEnumShellItems, BHID_EnumItems, IShellItem2, PropertiesSystem::{PROPERTYKEY, PropVariantChangeType, PROPVAR_CHANGE_FLAGS, PVCHF_DEFAULT, PropVariantToBSTR}, SHCOLUMNINFO,
        }, Globalization::lstrlenW,
    },
};

// Trying to convert this old Raymond Chen article to Rust: https://devblogs.microsoft.com/oldnewthing/20110831-00/?p=9763
// attempting to list items in the Recycle Bin
fn main() -> Result<()> {
    unsafe {
        CoInitializeEx(std::ptr::null(), Com::COINIT_APARTMENTTHREADED)?;

        // Couldn't find these in the `windows` crate
        // {43826D1E-E718-42EE-BC55-A1E261C37BFE}
        let IID_IShellItem = GUID::from_values(
            0x43826d1e,
            0xe718,
            0x42ee,
            [0xbc, 0x55, 0xa1, 0xe2, 0x61, 0xc3, 0x7b, 0xfe],
        );

        // {9B174B33-40FF-11d2-A27E-00C04FC30871}
        let PSGUID_DISPLACED = GUID::from_values(0x9b174b33, 0x40ff, 0x11d2, [0xa2, 0x7e, 0x0, 0xc0, 0x4f, 0xc3, 0x8, 0x71]);
        let PID_DISPLACED_FROM: u32 = 2;

        // TODO: make sure we free this
        let mut recycle_bin = MaybeUninit::<Option<IShellItem2>>::uninit();

        SHGetKnownFolderItem(
            &FOLDERID_RecycleBinFolder,
            KF_FLAG_DEFAULT,
            HANDLE::default(),
            &IID_IShellItem,
            recycle_bin.as_mut_ptr() as *mut *mut c_void,
        )?;

        let recycle_bin = recycle_bin.assume_init().expect("not initialized");
        
        let name = recycle_bin.GetDisplayName(SIGDN_NORMALDISPLAY)?;
        let path_copy = convert_to_os_string(name);
        dbg!(path_copy);

        let pesi: IEnumShellItems = recycle_bin.BindToHandler(None, &BHID_EnumItems)?;
        let mut fetched: u32 = 0;

        loop {
            // Not sure if this is the best way to get data out of Next
            let mut arr= [None];
            pesi.Next(&mut arr, &mut fetched)?;

            if fetched == 0 {
                break;
            }

            match &arr[0] {
                Some(i) => {
                    println!("{:?}", get_display_name(i)?);
                    let i2: IShellItem2 = i.cast()?;
                    // i2.GetPropertyDescriptionList()

                    // {9B174B33-40FF-11d2-A27E-00C04FC30871}
// #define PSGUID_DISPLACED    {0x9b174b33, 0x40ff, 0x11d2, 0xa2, 0x7e, 0x0, 0xc0, 0x4f, 0xc3, 0x8, 0x71}
// DEFINE_GUID(FMTID_Displaced, 0x9b174b33, 0x40ff, 0x11d2, 0xa2, 0x7e, 0x0, 0xc0, 0x4f, 0xc3, 0x8, 0x71);
// #define PID_DISPLACED_FROM  2
                    let l = PROPERTYKEY{ fmtid: PSGUID_DISPLACED, pid: PID_DISPLACED_FROM };
                    // let l: SHCOLUMNINFO;
                    let asdf = i2.GetProperty(&l)?;


                    let b = PropVariantToBSTR(&asdf)?;
                    // PropVariantChangeType()
                    dbg!(b);

                    // asdf.try_into()

                    // let a = &asdf.Anonymous.Anonymous.Anonymous.bstrVal;

                    // &asdf.
                    // WHY IS THIS CRASHING?
                    // let sadfasdgf = OsString::from_wide(a.as_wide());
                    // dbg!();
                },
                None => {break;},
            }
        }
    }

    println!("Done!");
    Ok(())
}

fn get_display_name(psi: &IShellItem) -> Result<OsString> {
    unsafe {
        let name = psi.GetDisplayName(SIGDN_NORMALDISPLAY)?;
        let path_copy = convert_to_os_string(name);
        // dbg!(path_copy);

        let ret = path_copy.clone();

        CoTaskMemFree(name.0 as *const c_void);
        Ok(ret)
    }
}

// void PrintDisplayName(IShellItem *psi, SIGDN sigdn, PCTSTR pszLabel)
// {
//  LPWSTR pszName;
//  HRESULT hr = psi->GetDisplayName(sigdn, &pszName);
//  if (SUCCEEDED(hr)) {
//   _tprintf(TEXT("%s = %ws\n"), pszLabel, pszName);
//   CoTaskMemFree(pszName);
//  }
// }

// no idea why this is so hard
fn convert_to_os_string(name: PWSTR) -> OsString {
    let converted = PCWSTR(name.0);
    // why is this conversion needed?
    unsafe {
        let path_ref: &[u16] = slice::from_raw_parts(name.0, lstrlenW(converted) as usize);
        OsString::from_wide(path_ref)
    }
}
