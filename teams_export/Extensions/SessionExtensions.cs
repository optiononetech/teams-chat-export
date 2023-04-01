using Microsoft.Graph;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Web;

public static class HttpSessionExtensions
{
    public static ConcurrentDictionary<string, string> Tracker = new ConcurrentDictionary<string, string>();
    public static ConcurrentDictionary<string, List<object>> Result = new ConcurrentDictionary<string, List<object>>();
    public static ConcurrentDictionary<string, List<object>> Delta = new ConcurrentDictionary<string, List<object>>();

    public static void SetCurrentAction(this HttpContextBase session, string actionKey, string actionValue)
    {
        Tracker[actionKey] = actionValue;
    }

    public static string GetCurrentAction(this HttpContextBase session, string actionKey)
    {
        if (Tracker.TryGetValue(actionKey, out string actionValue)) return actionValue;
        return "";
    }

    public static void SetResult<T>(this HttpContextBase session, string actionKey, List<T> actionValue)
    {
        Result[actionKey] = actionValue.Select(p => (object)p).ToList();
    }

    public static List<object> GetDelta(this HttpContextBase session, string actionKey)
    {
        List<object> result = new List<object>();
        if (Result.TryGetValue(actionKey, out List<object> currentSet))
        {
            if (Delta.TryGetValue(actionKey, out List<object> deltaSet))
                result = currentSet.Except(deltaSet).ToList();
            else
                result = currentSet;
            Delta[actionKey] = currentSet;
        }
        return result;
    }
}