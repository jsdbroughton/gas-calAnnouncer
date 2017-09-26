// Forces the authorization check, sometimes advanced services slip through the net.
var cal = Calendar, doc = DocumentApp, cc = CalendarApp;

function onOpen() {
  var d = DocumentApp.getUi()
  .createMenu('Schedule Announcer')
  .addItem('Update Announcements', 'runTime')
  .addItem('Calendar Settings', 'setCal')
  .addToUi();
}

function setCal() {
  var d = DocumentApp.getUi().showSidebar(HtmlService.createTemplateFromFile('sidebar')
      .evaluate().setTitle('Synchronisation Settings'));
}

function setCalProperties(settings) {
  
  var cal = settings.cals;
  var detail = settings.detail;
  
  var p = PropertiesService.getDocumentProperties().setProperty('calId', cal);
  var d = PropertiesService.getDocumentProperties().setProperty('detail', detail);
}

function getCalendars() {
  var cals = CalendarApp.getAllCalendars();
  var p = PropertiesService.getDocumentProperties().getProperty('calId');
  var d = PropertiesService.getDocumentProperties().getProperty('detail');
  
  return {
    detail: d,
    cals: cals.map(function(cal) { // cal = CalendarApp.getAllCalendars()[0];
      var c = {};
      
      c.name = cal.getName();
      c.id = cal.getId();
      c.selected = (c.id == p);
      
    return c;   
    })
  };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function runTime(){  
  
  var calId = '', det = 'false';
  var p = PropertiesService.getDocumentProperties()
    
  try {
    calId = p.getProperty('calId');
    det = p.getProperty('detail');
    Announcer.setDetail(det);
    Announcer.setCalendar(calId);
  } catch(err) {
    try {
      Announcer.setCalendar('stardotbmp.com_u8i4a0s5ke16a28ikftj8kunqs@group.calendar.google.com');
      calId = Announcer.getCalendar().getId();
    } catch (err) {
      DocumentApp.getUi().alert(err);
    }
  }
  
  if (calId) { 
    var d = DocumentApp.getActiveDocument()
    var events = Announcer.loadEvents().events;  
    var script = Announcer.writeScript(d.getId());
    
  }
  
  var dd = Announcer.detail;
  
  debugger;
}

var Announcer = (function(announcer) {
  
  Object.defineProperties(announcer, {
    'configuration': {
      value: {lookAhead: 10}
    }
  })
  
  Object.defineProperty(announcer, 'events', {
    get: function() {return eventsStore}
  });
  
  Object.defineProperty(announcer, 'scriptId', {
    value: '',
    writable: true,
    enumerable: true
  });
  
  var calendarId,
      Events = function(){
        return {
          today: [],
          tomorrow: [],
          thisWeek: [],
          nextWeek: [],
          thisWeekend: []};
      },
      eventsStore,
      eventsLoaded,
      detail = 'false';
  
  // All the date related calcs  
  var today = new Date(new Date().setHours(0,0,0,0)),
      endTime = new Date((new Date()).getTime() + 1000 * 60 * 60 * 24 * announcer.configuration.lookAhead), // gathers for the coming fortnight plus weekend.
      dayToday = today.getDay(),
      weekendShift = (24*60*60*1000*2),
      todayAdjust = 0,
      tomorrowAdjust = 0;
  
  if (dayToday == 5) { 
    todayAdjust = weekendShift;
    tomorrowAdjust = weekendShift;
  }
  
  if (dayToday == 4) {
    tomorrowAdjust = weekendShift;
  }
  
  var getWeek = function(date) {
    
    var year = date.getFullYear();
    var onejan = new Date(year, 0, 1);
    var week = Math.ceil((((date - onejan) / 86400000) + onejan.getDay() + 1) / 7);
    return week;
  };
  
  var tomorrowToday = (today.setHours(0,0,0,0) + 1000*60*60*24 + todayAdjust),
      tomorrowTomorrow = (today.setHours(0,0,0,0) + 1000*60*60*24*2 + tomorrowAdjust),
      tomorrowDay = (new Date(tomorrowTomorrow)).getDay();
  thisWeek = getWeek(new Date(tomorrowToday));
  
  var dayInWords = (function(d) {
    
    var m_names = new Array("January", "February", "March",
                            "April", "May", "June",
                            "July", "August", "September",
                            "October", "November", "December");
    
    var curr_date = d.getDate();
    var sup = "th";
    if (curr_date == 1 || curr_date == 21 || curr_date ==31) { sup = "st"; } 
    else if (curr_date == 2 || curr_date == 22){ sup = "nd"; } 
    else if (curr_date == 3 || curr_date == 23) { sup = "rd"; }
    
    var curr_month = d.getMonth();
    var curr_year = d.getFullYear();
    
    return [curr_date + sup, m_names[curr_month], + curr_year].join(' ');
  }((new Date(tomorrowToday))));
  
  void(0);
  
  announcer.setScript = function(id) {
    try {
      var d = DocumentApp.openById(scriptId);
      announcer.scriptId = scriptId;
    } catch(err) {
      announcer.scriptId = '';
    }
    return announcer;
  }
  
  announcer.setDetail = function(d) {
    announcer.detail = d;
  }
  
  announcer.writeScript = function(scriptId) {  
    
    if (calendarId && eventsLoaded) {
      
      scriptId = scriptId || announcer.scriptId;
      var name = 'Announcer Script for ' + dayInWords;
      var d;
      
      try {
        d = DocumentApp.getActiveDocument();
        d.setName(name);
      } catch (err) {
        DocumentApp.getUi().alert("No document found.");
        d = DocumentApp.create(name);
        announcer.scriptId = d.getId();
      }    
      
      var body = d.getBody().clear();
      
      body.appendParagraph(name).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendHorizontalRule();
      
      var today = eventsStore.today.slice(0);
      var tomorrow = eventsStore.tomorrow.slice(0);
      var thisWeekend = eventsStore.thisWeekend.slice(0);
      var thisWeek = eventsStore.thisWeek.slice(0);
      var nextWeek = eventsStore.nextWeek.slice(0);
      
      void(0);
      
      appendEvents(body, "Today", today);
      appendEvents(body, "Tomorrow", tomorrow);
      appendEvents(body, "This Week", thisWeek);
      appendEvents(body, "This Weekend", thisWeekend);      
      appendEvents(body, "Next Week", nextWeek);
      
      d.getBody().getChild(0).removeFromParent();
      
      d.saveAndClose();
      
      void(0);
    }
  }
  
  var appendEvents = function(body, header, events, detail) {
    
    if(events.length > 0) {
      
      headerBold = {};
      headerBold[DocumentApp.Attribute.BOLD] = true;
      
      var header = body.appendParagraph(header).setHeading(DocumentApp.ParagraphHeading.HEADING2).setSpacingAfter(0).setSpacingBefore(10).setAttributes(headerBold);
      
      events.forEach(function(item){
        
        var title = item.summary,
            description= (item.description || ''),
            time = (item.allDay ? 'All Day' : '⏰\t' + item.time),
            location = (item.location ? '@\t' + item.location : '\tNo Location Given');
        
        var italic = {};
        italic[DocumentApp.Attribute.ITALIC] = true;
        
        var normal = {};
        normal[DocumentApp.Attribute.ITALIC] = false;
        
        item.titleHeader = body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING3);
        item.descriptionPara = body.appendParagraph(description).setAttributes(italic).setSpacingAfter(8);
        
        if (Announcer.detail !== "false") {
          item.timePara = body.appendParagraph(time).setAttributes(normal);
        }
        
        if (item.location) {
          item.locationPara = body.appendParagraph(location).setAttributes(normal);
        }
        
        void(0)
        
      });
      
      var rule = body.appendHorizontalRule();
    }
  }
  
  announcer.setCalendar = function(id) {
    try {
      var c = Calendar.Calendars.get(id);
      if (c) {  
        calendarId = c.id;
      }
    } catch(err) {
      calendarId = undefined;
    }
    return this;
  }
  
  announcer.getCalendar = function() {
    try {
      return Calendar.Calendars.get(calendarId);
    } catch(err) {
      return undefined;
    }
  }
  
  announcer.loadEvents = function() {
    var events = new Events();
    
    if (calendarId) {
      // Query the calendar
      Calendar.Events.list(calendarId, {
        timeMin: (new Date(tomorrowToday)).toISOString(),
        timeMax: endTime.toISOString()
      })
      .items
      .forEach(function(evt){
        var event = {};
        
        event.location = evt.location || '';    
        event.description = evt.description || '';
        event.summary = evt.summary || '';
        event.time = evt.start.dateTime || (new Date(evt.start.date)).toISOString();
        event.allDay = !evt.start.dateTime;
        event.recurrence = (evt.recurrence || []).map(function(rec) { // rec = evt.recurrence[0]
          var r = {},
              freq = rec.match(/FREQ=(.*?);/),
              day = rec.match(/BYDAY=([A-Z]{2},{0,1})*?/g);
          
          if (freq) r.freq = freq ? freq[1] : undefined;
          if (day) r.day = day ? day[0].split('=')[1].split(',') : undefined;          
          return rec ? r : '';
        })[0];
        
        var eDate = (new Date(event.time)).setHours(0,0,0,0);
        
        var eventWeek = getWeek((new Date(eDate))),
            eDay = (new Date(eDate)).getDay();
        
        if (eDate == tomorrowToday) {
          Object.defineProperty(events, 'today', {
            enumerable: true
          });
          events.today.push(event);
          return;
        }
        
        if (eDate == tomorrowTomorrow && tomorrowDay == dayToday + 1) {
          events.tomorrow.push(event);
          return;
        }
        
        var d = (eDate - tomorrowToday) / (1000 * 60 * 60 * 24);
        
        if (thisWeek == eventWeek && eDay > dayToday || eDay == 0 && d < 7) {  
          if (eDay == 0 || eDay == 6) {
            events.thisWeekend.push(event);
          } else {
            events.thisWeek.push(event);
          }
          return;
        }
        
        if (eventWeek == thisWeek + 1 || thisWeek == eventWeek && eDay <= dayToday) {
          events.nextWeek.push(event);
          return;
        }    
        // discard any events not in scope.
      });
      
      // sort events
      for (list in events) {
        events[list] = events[list].sort(function(a, b) {
          var aD = new Date(a.time).getTime(), bD = new Date(b.time).getTime();
          return aD == bD ? 0 : aD < bD ? -1 : 1;
        });
      }
      eventsLoaded = true;
    } else {
      eventsLoaded = false;
    }
    eventsStore = events;
    return announcer;
  }
  return announcer;
}(Announcer || {}));