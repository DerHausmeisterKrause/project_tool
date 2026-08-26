using TaskTool.Infrastructure;
using TaskTool.Models;

namespace TaskTool.ViewModels;

public sealed class WebShortcutEditorViewModel : ObservableObject
{
    private string _id=Guid.NewGuid().ToString(),_name="Neue Webseite",_url="",_username="",_password=""; private bool _enabled=true,_autoLogin,_disableWebSecurity; private int _sortOrder;
    public string Id { get=>_id; set=>Set(ref _id,value); } public string Name { get=>_name; set=>Set(ref _name,value); } public string Url { get=>_url; set=>Set(ref _url,value); } public string Username { get=>_username; set=>Set(ref _username,value); } public string Password { get=>_password; set=>Set(ref _password,value); } public bool Enabled { get=>_enabled; set=>Set(ref _enabled,value); } public bool AutoLogin { get=>_autoLogin; set=>Set(ref _autoLogin,value); } public bool DisableWebSecurity { get=>_disableWebSecurity; set=>Set(ref _disableWebSecurity,value); } public int SortOrder { get=>_sortOrder; set=>Set(ref _sortOrder,value); }
    public string PasswordEncrypted { get; private set; } = "";
    public static WebShortcutEditorViewModel From(WebShortcutSettings x,string mask)=>new(){Id=x.Id,Name=x.Name,Url=x.Url,Username=x.Username,Password=x.PasswordEncrypted.Length>0?mask:"",PasswordEncrypted=x.PasswordEncrypted,Enabled=x.Enabled,AutoLogin=x.AutoLogin,DisableWebSecurity=x.DisableWebSecurity,SortOrder=x.SortOrder};
    public WebShortcutSettings ToModel()=>new(){Id=Id,Name=Name.Trim(),Url=Url.Trim(),Username=Username.Trim(),PasswordEncrypted=PasswordEncrypted,Enabled=Enabled,AutoLogin=AutoLogin,DisableWebSecurity=DisableWebSecurity,SortOrder=SortOrder}; public void SetEncrypted(string value)=>PasswordEncrypted=value;
}
