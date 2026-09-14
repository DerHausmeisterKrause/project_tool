using Microsoft.Data.Sqlite;
using TaskTool.Models;

namespace TaskTool.Services;

/// Imports Outlook snapshots only; it deliberately never calls Znuny or writes Outlook.
public sealed class PlenaroShareImportService
{
    private readonly DatabaseService _db; private readonly TaskService _tasks; private readonly SettingsService _settings; private readonly LoggerService _logger;
    public PlenaroShareImportService(DatabaseService db, TaskService tasks, SettingsService settings, LoggerService logger) { _db=db; _tasks=tasks; _settings=settings; _logger=logger; }
    public void Import(IEnumerable<OutlookCalendarEvent> events)
    {
        foreach(var e in events) try {
            if(!PlenaroShareCodec.TryParse(e.FullBody,out var p,out var hash,out var reason)){if(e.FullBody.Contains(PlenaroShareCodec.BeginMarker,StringComparison.Ordinal))_logger.Warning($"[PlenaroShareImport] action=skipped reason={reason}");continue;}
            if(string.Equals(p!.OriginClientInstanceId,_settings.Current.ClientInstanceId,StringComparison.OrdinalIgnoreCase)){_logger.Info("[PlenaroShareImport] action=skipped reason=self-origin");continue;}
            ImportOne(p,hash);
        } catch(Exception ex){_logger.Warning($"[PlenaroShareImport] action=skipped reason=event-error type={ex.GetType().Name}");}
    }
    private void ImportOne(PlenaroSharePayloadV1 p,string hash)
    {
        using var conn=new SqliteConnection(_db.ConnectionString);conn.Open(); Guid localId=Guid.Empty;string oldHash="";
        using(var c=conn.CreateCommand()){c.CommandText="SELECT local_task_id,last_payload_hash FROM plenaro_shared_task_imports WHERE task_share_id=$id";c.Parameters.AddWithValue("$id",p.TaskShareId);using var r=c.ExecuteReader();if(r.Read()){localId=Guid.Parse(r.GetString(0));oldHash=r.GetString(1);}}
        if(oldHash==hash)return; var all=_tasks.GetAllTasks(); TaskItem? task=localId==Guid.Empty?null:all.SingleOrDefault(x=>x.Id==localId);
        if(task==null&&!string.IsNullOrWhiteSpace(p.TicketId)){var matches=all.Where(x=>HasTicketId(x.Tags,p.TicketId)).ToList();if(matches.Count==1)task=matches[0];}
        var created=task==null;task??=new TaskItem{Status=TaskTool.Models.TaskStatus.Planned,IsZnunyAssigned=false};
        task.Title=p.TaskTitle;task.Description=p.TaskDescription;task.TicketUrl=p.TicketUrl;task.TicketState=p.TicketState;task.IsPlenaroShared=true;task.TaskShareId=p.TaskShareId;task.ShareOriginClientInstanceId=p.OriginClientInstanceId;
        if(!string.IsNullOrWhiteSpace(p.TicketId))task.Tags=MergeTicketTags(task.Tags,p.TicketId,p.TicketNumber);if(created)_tasks.CreateTask(task,true);else _tasks.UpdateTask(task,false);
        using(var c=conn.CreateCommand()){c.CommandText="INSERT INTO plenaro_shared_task_imports(task_share_id,local_task_id,origin_client_instance_id,last_payload_hash) VALUES($s,$l,$o,$h) ON CONFLICT(task_share_id) DO UPDATE SET local_task_id=$l,origin_client_instance_id=$o,last_payload_hash=$h";c.Parameters.AddWithValue("$s",p.TaskShareId);c.Parameters.AddWithValue("$l",task.Id.ToString());c.Parameters.AddWithValue("$o",p.OriginClientInstanceId);c.Parameters.AddWithValue("$h",hash);c.ExecuteNonQuery();}
        TaskSegment? segment=null;using(var c=conn.CreateCommand()){c.CommandText="SELECT local_segment_id FROM plenaro_shared_segment_imports WHERE segment_share_id=$id";c.Parameters.AddWithValue("$id",p.SegmentShareId);var id=c.ExecuteScalar();if(id!=null)segment=_tasks.GetSegments(task.Id).SingleOrDefault(x=>x.Id==Convert.ToInt64(id));}
        segment??=new TaskSegment{TaskId=task.Id,SegmentShareId=p.SegmentShareId,IsSharedImport=true};segment.StartLocal=p.SegmentStartLocal;segment.EndLocal=p.SegmentEndLocal;segment.Note=p.SegmentNote;segment.PlannedMinutes=(int)(segment.EndLocal-segment.StartLocal).TotalMinutes;if(segment.Id==0)_tasks.AddSegment(segment);else _tasks.UpdateSegment(segment);
        using(var c=conn.CreateCommand()){c.CommandText="INSERT INTO plenaro_shared_segment_imports(segment_share_id,local_segment_id,task_share_id,last_payload_hash) VALUES($s,$l,$t,$h) ON CONFLICT(segment_share_id) DO UPDATE SET local_segment_id=$l,last_payload_hash=$h";c.Parameters.AddWithValue("$s",p.SegmentShareId);c.Parameters.AddWithValue("$l",segment.Id);c.Parameters.AddWithValue("$t",p.TaskShareId);c.Parameters.AddWithValue("$h",hash);c.ExecuteNonQuery();}
        _logger.Info($"[PlenaroShareImport] action={(created?"import-created":"import-updated")} localTaskId={task.Id} hasTicket={!string.IsNullOrWhiteSpace(p.TicketId)}");
    }
    private static bool HasTicketId(string tags,string id)=>tags.Split(';',StringSplitOptions.TrimEntries).Any(x=>string.Equals(x,$"ZnunyTicketID:{id}",StringComparison.OrdinalIgnoreCase));
    private static string MergeTicketTags(string tags,string id,string number){var v=tags.Split(';',StringSplitOptions.RemoveEmptyEntries|StringSplitOptions.TrimEntries).Where(x=>!x.StartsWith("ZnunyTicketID:",StringComparison.OrdinalIgnoreCase)&&!x.StartsWith("ZnunyTicketNumber:",StringComparison.OrdinalIgnoreCase)).ToList();v.Add($"ZnunyTicketID:{id}");if(!string.IsNullOrWhiteSpace(number))v.Add($"ZnunyTicketNumber:{number}");return string.Join(';',v);}
}
