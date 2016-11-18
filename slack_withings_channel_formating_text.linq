<Query Kind="Program">
  <Reference>&lt;RuntimeDirectory&gt;\System.Threading.dll</Reference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Net.Http</NuGetReference>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

//Common Settings
public enum Member { dosyun, matzz, izak, shibayan, Kimuny }
public enum Month { January = 1, February = 2, March = 3, April = 4, May = 5, June = 6, July = 7, August = 8, September = 9, October = 10, November = 11, December = 12 }

//Slack Settings
public const string SlackHistoryApiEndpointUrl = "https://slack.com/api/groups.history";
public const string SlackToken = "";
public const string SlackChannelId = "";
public const int GetDataCount = 1000;
public const int AggregatePeriod = 300;

//SpreadSheet Settings
public const string SpreadSheetApiEndpointUrl = "https://sheets.googleapis.com/v4/spreadsheets/{0}:batchUpdate";
public const string TargetSpreadSheetId = "";
public const string GoogleOAuthAccessToken = "";

void Main()
{
	var url = string.Format("{0}?token={1}&channel={2}&count={3}", SlackHistoryApiEndpointUrl, SlackToken, SlackChannelId, GetDataCount);
	var json = SendRequest(url);
	var createMemberMeasureResultList = CreateMeasureResult(json);
	var text = CreateText(createMemberMeasureResultList).Dump();
	WriteSpreadSheet(createMemberMeasureResultList);

}

private string SendRequest(string url)
{
	var wc = new System.Net.WebClient();
	var result = wc.DownloadData(new Uri(url));
	return Encoding.UTF8.GetString(result);
}

public MeasureResult[] CreateMeasureResult(string json)
{
	var deserialzeObject = JsonConvert.DeserializeObject<SlackResponse>(json);
	return deserialzeObject.messages.Where(x => x.attachments != null)
										.Select(y => new { Title = y.attachments.First().title, Text = y.attachments.First().text })
										.Where(z => z.Text != null && z.Text.Contains("体重") && !z.Title.Contains("Amazon") && !z.Title.Contains("ダイエット"))
										.OrderBy(x => x.Title)
										.Select(x =>
										{
											var member = GetMember(x.Title);
											var baseParseText = DeleteMemberName(x.Title);
											var monthDeleteText = DeleteMonth(baseParseText);
											var year = GetYear(monthDeleteText);
											var month = GetMonth(baseParseText);
											var day = GetDay(monthDeleteText);
											var hour = GetHour(monthDeleteText);
											var minute = GetMinute(monthDeleteText);
											var weight = GetWeight(x.Text, member);
											var fatRate = GetFatRate(x.Text, member);

											return new MeasureResult() { Member = member, MeasureDate = new DateTime(year, month, day, hour, minute, 0), Weight = weight, FatRate = fatRate };
										})
										.Where(x => x.MeasureDate <= DateTime.Now && x.MeasureDate >= DateTime.Now.AddDays(AggregatePeriod * -1))
										.OrderByDescending(x => x.Member)
										.ThenBy(x => x.MeasureDate)
										.ToArray();


}


public Member GetMember(string name)
{
	if (name.Contains(Member.dosyun.ToString()))
		return Member.dosyun;
	else if (name.Contains(Member.shibayan.ToString()))
		return Member.shibayan;
	else if (name.Contains(Member.izak.ToString()))
		return Member.izak;
	else if (name.Contains(Member.matzz.ToString()))
		return Member.matzz;
	else if (name.Contains(Member.Kimuny.ToString()))
		return Member.Kimuny;

	return Member.dosyun;
}

public string DeleteMemberName(string title)
{
	var members = Enum.GetValues(typeof(Member));
	foreach (var member in members)
	{
		title = title.Replace(member.ToString(), string.Empty);
	}
	return title.Replace("'s", string.Empty).Trim();
}

public int GetMonth(string title)
{
	var monthList = Enum.GetValues(typeof(Month));
	foreach (var month in monthList)
	{
		if (title.Contains(month.ToString()))
			return (int)month;
	}
	return 0;
}

public string DeleteMonth(string title)
{
	var monthList = Enum.GetValues(typeof(Month));
	foreach (var month in monthList)
	{
		title = title.Replace(month.ToString(), string.Empty);
	}
	return title.Trim();
}

public int GetDay(string deleteMonthText)
{
	return int.Parse(deleteMonthText.Substring(0, 2));
}

public int GetYear(string deleteMonthText)
{	
	return int.Parse(deleteMonthText.Substring(3, 4));
}

public int GetHour(string deleteMonthText)
{
	var hour = int.Parse(deleteMonthText.Substring(11, 2));

	if (deleteMonthText.Contains("PM"))
		hour = hour + 12;

	if (hour == 24)
		return 0;

	return hour;
}

public int GetMinute(string deleteMonthText)
{
	return int.Parse(deleteMonthText.Substring(14, 2));
}

public double GetWeight(string text, Member memberName)
{
	if (memberName == Member.dosyun)
		return double.Parse(text.Substring(3, 4));
	else if (memberName == Member.shibayan)
		return double.Parse(text.Substring(4, 4));
	else if (memberName == Member.izak)
		return double.Parse(text.Substring(4, 4));
	else if (memberName == Member.matzz)
		return double.Parse(text.Substring(3, 6));
	else if (memberName == Member.Kimuny)
		return double.Parse(text.Substring(4, 4));
	return 0D;
}

public double GetFatRate(string text, Member memberName)
{
	if (text.Contains("--"))
		return 0D;
	else if (memberName == Member.dosyun)
		return double.Parse(text.Substring(17, 4));
	else if (memberName == Member.shibayan)
		return double.Parse(text.Substring(18, 4));
	else if (memberName == Member.izak)
		return double.Parse(text.Substring(18, 4));
	else if (memberName == Member.matzz)
		return double.Parse(text.Substring(18, 4));
	else if (memberName == Member.Kimuny)
		return double.Parse(text.Substring(18, 4));
	return 0D;
}

public string CreateText(MeasureResult[] result)
{
	var text = new StringBuilder();
	foreach (var member in Enum.GetValues(typeof(Member)))
	{
		var memberWeeklyResultList = result.Where(x => x.Member.ToString() == member.ToString())
										   .GroupBy(x => x.MeasureDate.Date)
										   .ToDictionary(x => x.Key, y => y.OrderBy(x => x.Weight).FirstOrDefault());

		text.Append(string.Format("◆{0}'s Weekly Result", member.ToString()));
		text.Append("\r\n");
		MeasureResult previousData = null;
		foreach (var currentData in memberWeeklyResultList)
		{
			text.Append(currentData.Key.ToString("yyyy/MM/dd(ddd)"));

			var currentValue = currentData.Value;

			var differenceWeight = 0D;
			var differenceFatRate = 0D;
			if (previousData != null)
			{
				differenceWeight = Math.Round(currentValue.Weight - previousData.Weight, 2);
				differenceFatRate = Math.Round(currentValue.FatRate - previousData.FatRate, 2);
			}
			text.Append(string.Format(" {0}kg ({1}{2}kg) ", currentValue.Weight, differenceWeight == 0D ? "±" : differenceWeight > 0 ? "+" : "", differenceWeight).PadRight(18));
			text.Append(string.Format(" {0}% ({1}{2}%)", currentValue.FatRate, differenceWeight == 0D ? "±" : differenceFatRate > 0 ? "+" : "", differenceFatRate).PadRight(18));
			previousData = currentData.Value;
			text.Append("\r\n");
		}

		var firstDayResult = memberWeeklyResultList.OrderBy(x => x.Value.MeasureDate).FirstOrDefault();
		var lastDayResult = memberWeeklyResultList.OrderByDescending(x => x.Value.MeasureDate).FirstOrDefault();

		if (firstDayResult.Value != null && lastDayResult.Value != null)
		{
			var differenceWeight = 0D;
			var differenceFatRate = 0D;
			differenceWeight = Math.Round(lastDayResult.Value.Weight - firstDayResult.Value.Weight, 2);
			differenceFatRate = Math.Round(lastDayResult.Value.FatRate - firstDayResult.Value.FatRate, 2);
			text.Append("\r\n");
			text.Append(string.Format("進捗 体重：{0}{1}kg 体脂肪率：{2}{3}%"
				   , differenceWeight == 0D ? "±" : differenceWeight > 0 ? "+" : ""
				   , differenceWeight
				   , differenceFatRate == 0D ? "±" : differenceFatRate > 0 ? "+" : ""
				   , differenceFatRate));


			text.Append("\r\n");
		}

		text.Append("------------------------------------------------");
		text.Append("\r\n");
	}

	return text.ToString();

}

public void WriteSpreadSheet(MeasureResult[] measureResultList)
{
	//valueを加算していくと列が増える
	//rowを加算していくと行が増える
	var rows = new List<Row>();
	var firstValues = new List<Value>();
	var firstRow = new Row();
	//左上はダミー行
	firstValues.Add(new Value());
	var memberlist = measureResultList.GroupBy(x => x.Member).ToArray();
	foreach (var member in memberlist)
	{
		//メンバーの名前を入れる
		var value = new Value() { userEnteredValue = new UserEnteredvalue() { stringValue = member.Key.ToString() } };
		firstValues.Add(value);
		firstRow.values = firstValues.ToArray();
	}
	//1行目終了
	rows.Add(firstRow);

	var measureDateList = measureResultList.GroupBy(x => x.MeasureDate.ToString("yyyy/MM/dd")).OrderBy(x => x.Key).ToArray();

	foreach (var date in measureDateList)
	{
		//日付を入れる
		var secondValues = new List<Value>();
		var value = new Value() { userEnteredValue = new UserEnteredvalue() { stringValue = date.Key } };
		secondValues.Add(value);

		foreach (var member in memberlist)
		{
			var memberMinWeightResult = member.Where(x => x.MeasureDate.ToString("yyyy/MM/dd") == date.Key).OrderBy(x => x.Weight).FirstOrDefault();
			if (memberMinWeightResult != null)
			{
				var weightValue = new Value() { userEnteredValue = new UserEnteredvalue() { stringValue = memberMinWeightResult.Weight.ToString() } };
				secondValues.Add(weightValue);
			}
			else
			{
				secondValues.Add(new Value());
			}
		}

		var secondRow = new Row();
		secondRow.values = secondValues.ToArray();
		rows.Add(secondRow);
	}
	
	var start = new Start() { rowIndex = 0, columnIndex = 0, sheetId = 0 };
	var updCells = new Updatecells() { start = start, rows = rows.ToArray(), fields = "userEnteredValue" };
	var req = new Request() { updateCells = updCells };
	var reqs = new Request[] { req };
	var root = new Rootobject() { requests = reqs };

	var json = JsonConvert.SerializeObject(root);
	var wc = new System.Net.WebClient();
	wc.Headers.Add("Authorization", "Bearer " + GoogleOAuthAccessToken);
	wc.Headers.Set("Content-Type", "application/json");
	wc.Encoding = Encoding.UTF8;

	var res = wc.UploadString(new Uri(string.Format(SpreadSheetApiEndpointUrl, TargetSpreadSheetId)), "POST", json);

	wc.Dispose();
}

public class MeasureResult
{
	public Member Member;
	public DateTime MeasureDate;
	public double Weight;
	public double FatRate;
}

//Slack Deserialize DTO
public class SlackResponse
{
	public string ok;
	public string has_more;
	public SlackMessage[] messages;
}

public class SlackMessage
{
	public string text;
	public SlackAttachment[] attachments;
}

public class SlackAttachment
{
	public string title;
	public string text;
}

//SpreadSheet Deserialize DTO
public class Rootobject
{
	public Request[] requests { get; set; }
}

public class Request
{
	public Updatecells updateCells { get; set; }
}

public class Updatecells
{
	public Start start { get; set; }
	public Row[] rows { get; set; }
	public string fields { get; set; }

}

public class Start
{
	public int sheetId { get; set; }
	public int rowIndex { get; set; }
	public int columnIndex { get; set; }
}

public class Row
{
	public Value[] values { get; set; }
}

public class Value
{
	public UserEnteredvalue userEnteredValue { get; set; }
}

public class UserEnteredvalue
{
	public string stringValue { get; set; }
}