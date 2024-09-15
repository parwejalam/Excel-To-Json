using System;

var allList = new List<dynamic>();

var chapterQueryList = new List<string>();
var sectionQueryList = new List<string>();
var questionQueryList = new List<string>();


var chapterGroup = allList.Group(x => x.ChapterId).ToList();
foreach(var chapter in chapterGroup) {

var chapterQuery = "";
chapterQueryList.Add(chapterQuery);

foreach (var section in chapter){
var sectionQuery = "";
sectionQueryList.Add(sectionQuery);

foreach (var question in section)

var questionQuery = "";
questionQueryList.Add(questionQuery);
}


}