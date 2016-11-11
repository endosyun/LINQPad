<Query Kind="Program">
  <Reference>&lt;RuntimeDirectory&gt;\System.Threading.dll</Reference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Net.Http</NuGetReference>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

//設定
public enum Member { dosyun, matzz, izak, shibayan, Kimuny }
public enum Month { January = 1, February = 2, March = 3, April = 4, May = 5, June = 6, July = 7, August = 8, September = 9, October = 10, November = 11, December = 12 }

public const string EndpointUrl = "https://slack.com/api/groups.history";
public const string Token = "";
public const string ChannelId =  "";
public const int GetDataCount = 100;

void Main()
{		
	var url = string.Format("{0}?token={1}&channel={2}&count={3}", EndpointUrl, Token, ChannelId, GetDataCount);
	sendRequest(url);
}

private async Task sendRequest(string url)
{
	var client = new HttpClient();
	var request = new HttpRequestMessage();

	request.Method = HttpMethod.Get;
	request.RequestUri = new Uri(url);

	var response = await client.SendAsync(request);
	string result = await response.Content.ReadAsStringAsync();

	var deserialzeObject = JsonConvert.DeserializeObject<SlackResponse>(result.ToString());
	var list = deserialzeObject.messages.Where(x => x.attachments != null)
										.Select(y => new { Title = y.attachments.First().title, Text = y.attachments.First().text })
										.Where(z => z.Text != null && z.Text.Contains("体重"))
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
										.Where(x => x.MeasureDate <= DateTime.Now && x.MeasureDate >= DateTime.Now.AddDays(-7))
										.OrderByDescending(x => x.Member)
										.ThenBy(x => x.MeasureDate)
										.ToArray();

	CreateText(list);
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
		title = title.Replace(member.ToString(),string.Empty);
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
		var memberWeeklyResultList = result.Where(x => x.Member.ToString() ==  member.ToString())
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

	return text.ToString().Dump();

}

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

public class MeasureResult
{
	public Member Member;
	public DateTime MeasureDate;
	public double Weight;
	public double FatRate;
}