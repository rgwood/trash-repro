use anyhow::Result;
use std::{
    ffi::{c_void, OsString},
    mem::MaybeUninit,
    os::windows::prelude::OsStringExt,
    slice,
};
use windows::{
    core::{Interface, GUID, PCWSTR, PWSTR},
    Win32::{
        Foundation::HANDLE,
        Globalization::lstrlenW,
        System::Com::CoInitializeEx,
        System::Com::{self, CoTaskMemFree},
        UI::Shell::{
            BHID_EnumItems, FOLDERID_RecycleBinFolder, IEnumShellItems, IShellItem, IShellItem2,
            PropertiesSystem::{PropVariantToBSTR, PROPERTYKEY},
            SHGetKnownFolderItem, KF_FLAG_DEFAULT, SIGDN_NORMALDISPLAY,
        },
    },
};

// Trying to convert this old Raymond Chen article to Rust: https://devblogs.microsoft.com/oldnewthing/20110831-00/?p=9763
// attempting to list items in the Recycle Bin
fn main() -> Result<()> {
    // Couldn't find these in the `windows` crate
    // {43826D1E-E718-42EE-BC55-A1E261C37BFE}
    #[allow(non_snake_case)]
    let IID_IShellItem = GUID::from_values(0x43826d1e, 0xe718, 0x42ee, [0xbc, 0x55, 0xa1, 0xe2, 0x61, 0xc3, 0x7b, 0xfe]);
    // {9B174B33-40FF-11d2-A27E-00C04FC30871}
    #[allow(non_snake_case)]
    let PSGUID_DISPLACED = GUID::from_values(0x9b174b33, 0x40ff, 0x11d2, [0xa2, 0x7e, 0x0, 0xc0, 0x4f, 0xc3, 0x8, 0x71]);
    #[allow(non_snake_case)]
    let PID_DISPLACED_FROM: u32 = 2;
    
    unsafe {
        CoInitializeEx(std::ptr::null(), Com::COINIT_APARTMENTTHREADED)?;

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
            let mut arr = [None];
            pesi.Next(&mut arr, &mut fetched)?;

            if fetched == 0 {
                break;
            }

            match &arr[0] {
                Some(item) => {
                    println!("{:?}", get_display_name(item)?);
                    let item2: IShellItem2 = item.cast()?;
                    // TODO: free this?
                    let original_location_variant = item2.GetProperty(
                        &(PROPERTYKEY {
                            fmtid: PSGUID_DISPLACED,
                            pid: PID_DISPLACED_FROM,
                        }),
                    )?;
                    let original_location = PropVariantToBSTR(&original_location_variant)?;
                    // PropVariantChangeType()
                    dbg!(original_location);
                }
                None => {
                    break;
                }
            }
        }
    }

    println!("Done!");
    Ok(())
}

/// Ported this C++:
/// void PrintDisplayName(IShellItem *psi, SIGDN sigdn, PCTSTR pszLabel)
/// {
///  LPWSTR pszName;
///  HRESULT hr = psi->GetDisplayName(sigdn, &pszName);
///  if (SUCCEEDED(hr)) {
///   _tprintf(TEXT("%s = %ws\n"), pszLabel, pszName);
///   CoTaskMemFree(pszName);
///  }
/// }
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


// no idea why this is so hard
fn convert_to_os_string(name: PWSTR) -> OsString {
    let converted = PCWSTR(name.0);
    // why is this conversion needed?
    unsafe {
        let path_ref: &[u16] = slice::from_raw_parts(name.0, lstrlenW(converted) as usize);
        OsString::from_wide(path_ref)
    }
}
