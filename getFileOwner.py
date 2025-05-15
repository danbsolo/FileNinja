import ctypes
from ctypes import wintypes
import threading

# CUSTOM
OWNER_INFO_CACHE = {}
CACHE_EVENTS = {}
CACHE_LOCK = threading.Lock()
dummyData = "DUMMY"

kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
advapi32 = ctypes.WinDLL('advapi32', use_last_error=True)

# Constants
SE_FILE_OBJECT = 1
OWNER_SECURITY_INFORMATION = 0x00000001

# Mapping for SID type values
sid_type_map = {
    1: 'User',
    2: 'Group',
    3: 'Domain',
    4: 'Alias',
    5: 'WellKnownGroup',
    6: 'DeletedAccount',
    7: 'Invalid',
    8: 'Unknown',
    9: 'Computer',
    10: 'Label'
}

# Set up GetNamedSecurityInfoW function prototype
advapi32.GetNamedSecurityInfoW.restype = wintypes.DWORD
advapi32.GetNamedSecurityInfoW.argtypes = [
    wintypes.LPWSTR,         # pObjectName
    wintypes.DWORD,          # ObjectType
    wintypes.DWORD,          # SecurityInfo (only OWNER requested)
    ctypes.POINTER(ctypes.c_void_p),  # ppsidOwner
    ctypes.POINTER(ctypes.c_void_p),  # ppsidGroup (unused)
    ctypes.POINTER(ctypes.c_void_p),  # ppDacl (unused)
    ctypes.POINTER(ctypes.c_void_p),  # ppSacl (unused)
    ctypes.POINTER(ctypes.c_void_p)   # ppSecurityDescriptor
]

# Set up LookupAccountSidW function prototype
advapi32.LookupAccountSidW.restype = wintypes.BOOL
advapi32.LookupAccountSidW.argtypes = [
    wintypes.LPCWSTR,         # lpSystemName (None for local)
    ctypes.c_void_p,          # Sid
    wintypes.LPWSTR,          # Name buffer
    ctypes.POINTER(wintypes.DWORD), # cchName
    wintypes.LPWSTR,          # Domain buffer
    ctypes.POINTER(wintypes.DWORD), # cchDomainName
    ctypes.POINTER(wintypes.DWORD)  # peUse (SID type)
]

def get_file_owner_info(filename):
    # Ensure the filename is a Unicode string.
    if not isinstance(filename, str):
        filename = filename.decode('utf-8')
    
    # Prepare pointers for owner SID and security descriptor.
    pOwner = ctypes.c_void_p()
    pSD = ctypes.c_void_p()
    
    error = advapi32.GetNamedSecurityInfoW(
        filename,
        SE_FILE_OBJECT,
        OWNER_SECURITY_INFORMATION,
        ctypes.byref(pOwner),
        None, None, None,
        ctypes.byref(pSD)
    )
    if error != 0:
        raise ctypes.WinError(error)
    
    # Prepare buffers for account name and domain.
    name_buf = ctypes.create_unicode_buffer(256)
    domain_buf = ctypes.create_unicode_buffer(256)
    name_size = wintypes.DWORD(256)
    domain_size = wintypes.DWORD(256)
    sid_type = wintypes.DWORD()
    
    if not advapi32.LookupAccountSidW(
        None, pOwner, name_buf, ctypes.byref(name_size),
        domain_buf, ctypes.byref(domain_size), ctypes.byref(sid_type)
    ):
        raise ctypes.WinError(ctypes.get_last_error())
    
    # Free the security descriptor memory.
    kernel32.LocalFree(pSD)
    
    sid_type_str = sid_type_map.get(sid_type.value, f"Unknown({sid_type.value})")
    return f"{domain_buf.value}\\{name_buf.value} ({sid_type_str})"

def getOwnerCatch(longFileAbsolute):
    """Return the owner info in 'DOMAIN\\Owner (SID_Type)' format. Return error info if applicable. Also manage OWNER_CACHE."""
    ownerInfo = OWNER_INFO_CACHE.get(longFileAbsolute)
    if ownerInfo is not None and ownerInfo != dummyData:
        return ownerInfo

    with CACHE_LOCK:
        ownerInfo = OWNER_INFO_CACHE.get(longFileAbsolute)

        if ownerInfo is None:
            OWNER_INFO_CACHE[longFileAbsolute] = dummyData    
            CACHE_EVENTS[longFileAbsolute] = threading.Event()
            computeOwner = True
        elif ownerInfo == dummyData:
            computeOwner = False
        else:
            return ownerInfo

    if computeOwner:
        try:
            ownerInfo = get_file_owner_info(longFileAbsolute)
        except Exception as e:
            ownerInfo = f"GET OWNER FAILED: {e}"
        
        OWNER_INFO_CACHE[longFileAbsolute] = ownerInfo
        CACHE_EVENTS[longFileAbsolute].set()
        return ownerInfo
    
    CACHE_EVENTS[longFileAbsolute].wait()
    return OWNER_INFO_CACHE[longFileAbsolute]

# Example usage:
if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        sys.exit(f"Usage: {sys.argv[0]} <filename>")
    filename = sys.argv[1]
    print("Result:", getOwnerCatch(filename))
