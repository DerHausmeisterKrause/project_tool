using Microsoft.Data.Sqlite;
using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class PlenaroShareTests
{
    [Fact]
    public void AttendeesAreValidatedAndDeduplicatedCaseInsensitively()
    {
        Assert.True(SegmentAttendeeParser.TryParse("Max@firma.de; max@FIRMA.de, anna@firma.de\nbob@firma.de", out var values, out _));
        Assert.Equal(3, values.Count);
        Assert.False(SegmentAttendeeParser.TryParse("valid@firma.de; xyz", out _, out var invalid));
        Assert.Equal("xyz", invalid);
    }

    [Fact]
    public void PayloadRoundTripsAndTamperingIsRejected()
    {
        var payload = Sample();
        var body = PlenaroShareCodec.AppendOrReplace("Beschreibung", payload, out var hash);
        Assert.True(PlenaroShareCodec.TryParse(body, out var parsed, out var parsedHash, out _));
        Assert.Equal(payload.TaskShareId, parsed!.TaskShareId);
        Assert.Equal(hash, parsedHash);
        Assert.False(PlenaroShareCodec.TryParse(body.Replace("Payload: ", "Payload: A"), out _, out _, out var reason));
        Assert.Contains(reason, new[] { "checksum-invalid", "json-invalid" });
    }

    [Fact]
    public void ExistingBlockIsReplacedAndHumanBodyRemains()
    {
        var first = PlenaroShareCodec.AppendOrReplace("Menschlich", Sample(), out _);
        var second = PlenaroShareCodec.AppendOrReplace(first, Sample(), out _);
        Assert.StartsWith("Menschlich", second);
        Assert.Equal(1, second.Split(PlenaroShareCodec.BeginMarker).Length - 1);
    }

    [Fact]
    public void UnknownVersionAndAmbiguousBlocksAreRejected()
    {
        var body = PlenaroShareCodec.AppendOrReplace("Text", Sample(), out _);
        Assert.False(PlenaroShareCodec.TryParse(body.Replace("Version: 1", "Version: 2"), out _, out _, out _));
        Assert.False(PlenaroShareCodec.TryParse(body + "\n" + body, out _, out _, out var reason));
        Assert.Equal("ambiguous-blocks", reason);
    }

    [Fact]
    public void V29MigrationIsAdditiveAndKeepsV28Data()
    {
        var path=Path.GetTempFileName();
        try { using(var c=new SqliteConnection($"Data Source={path}")){c.Open();var cmd=c.CreateCommand();cmd.CommandText="CREATE TABLE schema_version(version INTEGER NOT NULL); INSERT INTO schema_version VALUES(28); CREATE TABLE tasks(id TEXT PRIMARY KEY,title TEXT NOT NULL,description TEXT,ticket_url TEXT,start_local TEXT,end_local TEXT,status TEXT NOT NULL,priority INTEGER,tags TEXT,outlook_entry_id TEXT,ticket_minutes_booked INTEGER NOT NULL DEFAULT 0,ticket_seconds_booked INTEGER NOT NULL DEFAULT 0,is_pinned INTEGER NOT NULL DEFAULT 0,created_utc TEXT NOT NULL,updated_utc TEXT NOT NULL); INSERT INTO tasks VALUES('keep','Keep','','',NULL,NULL,'Planned',NULL,'','',0,0,0,'2020-01-01','2020-01-01');";cmd.ExecuteNonQuery();}
            new DatabaseService(new LoggerService(AppLogLevel.Error),path).Initialize();
            using var check=new SqliteConnection($"Data Source={path}");check.Open();using var cmd2=check.CreateCommand();cmd2.CommandText="SELECT (SELECT version FROM schema_version),(SELECT title FROM tasks WHERE id='keep'),(SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='task_segment_attendees')";using var r=cmd2.ExecuteReader();Assert.True(r.Read());Assert.Equal(29,r.GetInt32(0));Assert.Equal("Keep",r.GetString(1));Assert.Equal(1,r.GetInt32(2));
        } finally { File.Delete(path); }
    }

    [Fact]
    public void SharedUnassignedZnunyTaskRemainsVisible()
    { var task=new TaskItem{Tags="ZnunyTicketID:42",IsZnunyAssigned=false,IsPlenaroShared=true};Assert.True(task.IsOperationallyVisible); }

    private static PlenaroSharePayloadV1 Sample()=>new(){TaskShareId=Guid.NewGuid().ToString(),SegmentShareId=Guid.NewGuid().ToString(),OriginClientInstanceId=Guid.NewGuid().ToString(),TaskTitle="Titel",TaskDescription="Beschreibung",SegmentStartLocal=DateTime.Today.AddHours(9),SegmentEndLocal=DateTime.Today.AddHours(10),GeneratedUtc=DateTime.UtcNow};
}
