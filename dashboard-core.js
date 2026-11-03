/* Scopus Research Insights — browser/Node shared data model. */
(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ScopusDashboardCore = factory();
})(typeof self !== 'undefined' ? self : this, function() {
  'use strict';
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const QUARTILES = ['Q1','Q2','Q3','Q4','N/A'];
  const norm = v => String(v == null ? '' : v).trim().replace(/\s+/g,' ').toLowerCase();
  const keyify = v => norm(v).replace(/[^a-z0-9]/g,'');
  const value = (obj, ...aliases) => {
    if (!obj) return '';
    for (const name of aliases) {
      if (Object.prototype.hasOwnProperty.call(obj,name)) return obj[name] == null ? '' : String(obj[name]).trim();
    }
    const wanted = new Set(aliases.map(keyify));
    for (const k of Object.keys(obj)) if (wanted.has(keyify(k))) return obj[k] == null ? '' : String(obj[k]).trim();
    return '';
  };
  const number = v => {
    if (v === '' || v === undefined || v === null || v === 'N/A') return null;
    const n = Number(String(v).replace(/,/g,''));
    return Number.isFinite(n) ? n : null;
  };
  const validQuartile = v => { const x=String(v||'').trim().toUpperCase(); return /^Q[1-4]$/.test(x) ? x : 'N/A'; };
  const asMonth = (label,date) => {
    const m = MONTHS.find(x => norm(x) === norm(label));
    if (m) return m;
    const match = String(date||'').match(/^\d{4}-(\d{1,2})(?:-|$)/);
    const ix = match ? Number(match[1])-1 : -1;
    return ix>=0 && ix<12 ? MONTHS[ix] : 'Unknown';
  };
  // Canonical paper links are built from the Scopus DOCUMENT ID, not an author ID.
  // This also repairs older Excel reports whose URLs led to the XML API endpoint.
  function scopusRecordUrl(eid, originalUrl) {
    const raw = String(eid || '').trim();
    let paperId = (raw.match(/^2-s2\.0-(\d{6,20})$/i) || raw.match(/^(?:SCOPUS_ID\s*:\s*)?(\d{6,20})$/i) || [])[1] || '';
    const original = String(originalUrl || '').trim();
    if (!paperId && original) {
      try {
        const parsed = new URL(original);
        const host = parsed.hostname.toLowerCase();
        if (host === 'scopus.com' || host.endsWith('.scopus.com') || host === 'api.elsevier.com') {
          const idCandidate = parsed.searchParams.get('scp') || parsed.searchParams.get('eid') || '';
          paperId = (idCandidate.match(/^2-s2\.0-(\d{6,20})$/i) || idCandidate.match(/^(\d{6,20})$/) || [])[1] || '';
          if (!paperId) paperId = (parsed.pathname.match(/\/scopus_id\/(\d{6,20})(?:\/|$)/i) || [])[1] || '';
        }
      } catch (_) { /* broken or non-URL source link */ }
    }
    if (paperId) return `https://www.scopus.com/inward/record.uri?partnerID=HzOxMe3b&scp=${paperId}&origin=inward`;
    try {
      const url = new URL(original);
      if (url.protocol === 'https:' && ['www.scopus.com','scopus.com'].includes(url.hostname.toLowerCase())) return original;
    } catch (_) { /* no valid web page fallback */ }
    return '';
  }
  function scopusCitedByUrl(eid, originalUrl) {
    const r = scopusRecordUrl(eid, originalUrl);
    const match = r.match(/[?&]scp=(\d{6,20})(?:&|$)/);
    if (match) return `https://www.scopus.com/inward/citedby.url?partnerID=HzOxMe3b&scp=${match[1]}&origin=inward`;
    try {
      const u=new URL(String(originalUrl||''));
      if (u.protocol==='https:' && (u.hostname==='www.scopus.com'||u.hostname==='scopus.com') && u.pathname==='/inward/citedby.url') return u.toString();
    } catch (_) { }
    return '';
  }
  const yearOf = (...candidates) => {
    for (const v of candidates) {
      const m = String(v || '').match(/(?:^|\D)((?:19|20)\d{2})(?:\D|$)/);
      if (m) return +m[1];
    }
    return null;
  };
  function ingestReports(reports) {
    const authors = new Map(), records = new Map(), metrics = new Map(), errors = [];
    let missingRecordSheets = 0;
    const years = new Set();
    for (const report of reports) {
      const sheets = report.sheets || {};
      const summary = sheets['Monthly Summary'] || [];
      const rawRecords = sheets['Scopus Records'] || [];
      const rawMetrics = sheets['Journal Metrics'] || [];
      const runInfo = sheets['Run Info'] || [];
      const infoField = label => value((runInfo.find(x=>norm(value(x,'Report parameter'))===label) || {}),'Value');
      const fromYear = yearOf(infoField('from year'));
      const toYear = yearOf(infoField('to year'));
      const reportYear = yearOf(infoField('selected year'), report.name);
      if (fromYear && toYear && fromYear<=toYear && toYear-fromYear<=100) {
        for (let y=fromYear;y<=toYear;y++) years.add(y);
      }
      if (!rawRecords.length) missingRecordSheets++;
      function addAuthor(name,id,profile,scholar,status) {
        const cleanId = String(id||'').trim();
        const cleanName = String(name||'').trim();
        const authorKey = cleanId ? 'id:'+cleanId : (profile ? 'url:'+norm(profile) : 'name:'+norm(cleanName));
        if (!cleanName && !cleanId) return null;
        const existing = authors.get(authorKey);
        const updated = existing || {key:authorKey,name:cleanName||cleanId,id:cleanId,profile:profile||'',scholar:scholar||'',status:status||'',files:new Set(),reportedTotals:new Map(),statusByYear:new Map()};
        if (!updated.name || updated.name === updated.id) updated.name=cleanName||cleanId;
        updated.profile ||= profile||''; updated.scholar ||= scholar||'';
        if (status && norm(status) !== 'not checked') updated.status = status;
        updated.files.add(report.name);
        authors.set(authorKey,updated);
        return authorKey;
      }
      for (const row of summary) {
        if (norm(value(row,'Author Name')).startsWith('total (')) continue;
        const id=value(row,'Scopus Author ID');
        const name=value(row,'Author Name');
        const key=addAuthor(name,id,value(row,'Scopus Profile URL','Scopus Profile'),value(row,'Google Scholar Profile URL'),value(row,'Retrieval Status'));
        const count=number(value(row,'Annual Total'));
        const rowYear = yearOf(value(row,'Publication Year'), reportYear);
        if (rowYear) years.add(rowYear);
        if (key && rowYear) {
          if(count!==null) authors.get(key).reportedTotals.set(rowYear,count);
          const status=value(row,'Retrieval Status');
          if (status) authors.get(key).statusByYear.set(rowYear,status);
        }
      }
      for (const row of rawMetrics) {
        const name = value(row,'Journal / Source');
        if (!name) continue;
        const sid=value(row,'Source ID'); const issn=value(row,'ISSN','eISSN');
        const entry={name, sourceId:sid,issn,score:number(value(row,'CiteScore')),year:number(value(row,'CiteScore Year')),
          pct:number(value(row,'Highest CiteScore Percentile')),quartile:validQuartile(value(row,'Best CiteScore Quartile')),
          category:value(row,'Best Subject Category'),url:value(row,'Scopus Source URL'),status:value(row,'Retrieval Status')};
        for(const key of [sid?'id:'+sid:'',issn?'issn:'+norm(issn).replace(/-/g,''):'','name:'+norm(name)]) {
          if(key) {
            const old=metrics.get(key);
            if(!old || (entry.year||0)>=(old.year||0)) metrics.set(key,entry);
          }
        }
      }
      for (const row of rawRecords) {
        const name=value(row,'Author Name');
        const id=value(row,'Scopus Author ID');
        const profile=value(row,'Scopus Profile URL');
        const authorKey=addAuthor(name,id,profile,'','');
        if(!authorKey) continue;
        const title=value(row,'Paper Title','Title');
        const doi=value(row,'DOI').replace(/^https?:\/\/(?:dx\.)?doi\.org\//i,'');
        const eid=value(row,'Scopus EID');
        const date=value(row,'Publication Date');
        const year=yearOf(value(row,'Scopus Search Year'),date,reportYear);
        if(year) years.add(year);
        if(!title) continue;
        const paperKey=value(row,'Unique Paper Key') || (doi?'doi:'+norm(doi):eid?'eid:'+eid:'title:'+keyify(title)+':'+date);
        const uniqueRecordKey=authorKey+'|'+String(year)+'|'+paperKey;
        const entry={authorKey,authorName:authors.get(authorKey).name,title,doi,eid,date,year,month:asMonth(value(row,'Publication Month'),date),
          journal:value(row,'Journal / Source'),sourceId:value(row,'Scopus Source ID'),issn:value(row,'ISSN','eISSN'),
          type:value(row,'Source Type'),scopusUrl:scopusRecordUrl(eid,value(row,'Scopus Paper URL')),doiUrl:value(row,'DOI URL')||(doi?'https://doi.org/'+doi:''),
          paperKey,citationCount:number(value(row,'Scopus Citation Count','Citation Count','Cited By Count','citedby-count')),
          citedByUrl:scopusCitedByUrl(eid,value(row,'Scopus Cited By URL')),
          citationAsOf:value(row,'Citation Count As Of'),
          citeScore:number(value(row,'CiteScore')),citeScoreYear:number(value(row,'CiteScore Year')),
          percentile:number(value(row,'Highest CiteScore Percentile')),quartile:validQuartile(value(row,'Best CiteScore Quartile')),
          category:value(row,'Best Subject Category'),metricStatus:value(row,'Journal Metrics Status'),sourceUrl:value(row,'Scopus Source URL')};
        const existing=records.get(uniqueRecordKey);
        if(!existing) records.set(uniqueRecordKey,entry);
        else {
          // Prefer the latest citation snapshot, retaining known values over blank ones.
          const newer = entry.citationAsOf && (!existing.citationAsOf || entry.citationAsOf >= existing.citationAsOf);
          const betterCites = existing.citationCount===null && entry.citationCount!==null;
          if (entry.citationCount!==null && (betterCites || (newer && existing.citationCount!==null))) {
            existing.citationCount=entry.citationCount;
            existing.citationAsOf=entry.citationAsOf;
          }
          existing.citedByUrl ||= entry.citedByUrl;
          if (existing.quartile==='N/A' && entry.quartile!=='N/A') {
            existing.quartile=entry.quartile;existing.citeScore=entry.citeScore;
            existing.citeScoreYear=entry.citeScoreYear;existing.percentile=entry.percentile;
          }
        }
      }
      for(const row of (sheets.Errors||[])) if(value(row,'Issue')) errors.push({author:value(row,'Author'),stage:value(row,'Stage'),issue:value(row,'Issue'),file:report.name});
      if (reportYear) years.add(reportYear);
    }
    const papers=Array.from(records.values());
    for(const p of papers) {
      const m=metrics.get(p.sourceId?'id:'+p.sourceId:'') || metrics.get(p.issn?'issn:'+norm(p.issn).replace(/-/g,''):'') || metrics.get('name:'+norm(p.journal));
      if(m) {
        if(p.citeScore === null) p.citeScore=m.score;
        if(p.citeScoreYear === null) p.citeScoreYear=m.year;
        if(p.percentile === null) p.percentile=m.pct;
        if(p.quartile === 'N/A' && m.quartile!=='N/A') p.quartile=m.quartile;
        p.category ||= m.category; p.sourceUrl ||= m.url; p.metricStatus ||= m.status;
      }
    }
    const incompleteYears = new Set();
    for (const a of authors.values()) for (const [y,status] of a.statusByYear)
      if (/^(error|cancelled|incomplete)/i.test(status)) incompleteYears.add(y);
    for (const e of errors) {
      const year = yearOf(e.stage);
      if (year && /^scopus/i.test(e.stage)) incompleteYears.add(year);
    }
    return {authors:Array.from(authors.values()).sort((a,b)=>a.name.localeCompare(b.name)),records:papers,incompleteYears,
      years:Array.from(years).sort((a,b)=>b-a),journalMetrics:Array.from(new Map(Array.from(metrics.values()).map(v=>[v.sourceId||v.issn||v.name,v])).values()),
      errors,missingRecordSheets,reportCount:reports.length};
  }
  function matchingRecords(data,authorKey,yearOrRange) {
    const matchYear = y => {
      if (!yearOrRange) return true;
      if (typeof yearOrRange === 'number') return y===yearOrRange;
      const lo=Number(yearOrRange.from), hi=Number(yearOrRange.to);
      return y>=lo && y<=hi;
    };
    const matched=data.records.filter(p=>(!authorKey||p.authorKey===authorKey)&&matchYear(p.year));
    if(authorKey) return matched;
    const seen=new Map();
    for(const p of matched) {
      const key=String(p.year)+'|'+p.paperKey;
      const old=seen.get(key);
      if(!old) seen.set(key,{...p});
      else {
        // Coauthored paper counts are institutionally counted once. Reconcile
        // citation counts from the newest available snapshot of its author rows.
        if (p.citationCount!==null && (old.citationCount===null || (p.citationAsOf||'')>(old.citationAsOf||''))) {
          old.citationCount=p.citationCount;old.citationAsOf=p.citationAsOf;
        }
        old.citedByUrl ||= p.citedByUrl;
        if (old.quartile==='N/A' && p.quartile!=='N/A') old.quartile=p.quartile;
      }
    }
    return Array.from(seen.values());
  }
  function summarize(records) {
    const months=Object.fromEntries([...MONTHS,'Unknown'].map(m=>[m,0]));
    const quartiles=Object.fromEntries(QUARTILES.map(q=>[q,0]));
    const journals=new Set(); let scoreSum=0,scoreCount=0;
    let citationTotal=0,citationKnown=0,citationMissing=0,citedPapers=0,maxCitations=0;
    const knownCounts=[];
    for(const p of records) {
      months[p.month in months?p.month:'Unknown']++;
      quartiles[p.quartile in quartiles?p.quartile:'N/A']++;
      if(p.journal) journals.add(norm(p.journal));
      if(p.citeScore!==null) {scoreSum+=p.citeScore;scoreCount++;}
      if (p.citationCount!==null) {
        citationTotal+=p.citationCount;citationKnown++;
        if (p.citationCount>0) citedPapers++;
        maxCitations=Math.max(maxCitations,p.citationCount);
        knownCounts.push(p.citationCount);
      } else citationMissing++;
    }
    knownCounts.sort((a,b)=>b-a);
    const hIndex=knownCounts.reduce((acc,n,i)=>n>=i+1?i+1:acc,0);
    return {total:records.length,months,quartiles,journals:journals.size,meanCiteScore:scoreCount?scoreSum/scoreCount:null,
      classified:records.length-quartiles['N/A'],citationTotal,citationKnown,citationMissing,citedPapers,maxCitations,
      meanCitations:citationKnown?citationTotal/citationKnown:null,hIndex};
  }
  function rankAuthors(data,yearOrRange) {
    return data.authors.map(a=>({author:a,...summarize(matchingRecords(data,a.key,yearOrRange))}))
      .sort((a,b)=>b.total-a.total||b.quartiles.Q1-a.quartiles.Q1||a.author.name.localeCompare(b.author.name));
  }
  function compareYears(data, authorKey, from, to) {
    const result=[];
    for(let year=from;year<=to;year++) {
      const rows=matchingRecords(data,authorKey,year),s=summarize(rows);
      const involved=new Set(data.records.filter(p=>p.year===year).map(p=>p.authorKey));
      const person=authorKey?data.authors.find(x=>x.key===authorKey):null;
      const failedAuthor=person && /^(error|cancelled|incomplete)/i.test(person.statusByYear.get(year)||'');
      result.push({year,...s,authorsWithPapers:involved.size,
        incomplete:authorKey?Boolean(failedAuthor):data.incompleteYears.has(year)});
    }
    return result;
  }
  function compareAuthors(data,from,to) {
    return data.authors.map(author=>{
      const byYear={},citationsByYear={},citationMissingByYear={};let total=0,citationTotal=0;
      for(let year=from;year<=to;year++) {
        const status=author.statusByYear.get(year)||'';
        const isFailed=/^(error|cancelled|incomplete)/i.test(status);
        const count=isFailed?null:matchingRecords(data,author.key,year).length;
        byYear[year]=count;
        const s=summarize(matchingRecords(data,author.key,year));
        citationsByYear[year]=isFailed?null:s.citationKnown?s.citationTotal:(s.total===0?0:null);
        citationMissingByYear[year]=s.citationMissing;
        if(count!==null) total+=count;
        if(!isFailed) citationTotal+=s.citationTotal;
      }
      return {author,byYear,total,citationsByYear,citationMissingByYear,citationTotal};
    }).sort((a,b)=>b.total-a.total||a.author.name.localeCompare(b.author.name));
  }
  return {MONTHS,QUARTILES,ingestReports,matchingRecords,summarize,rankAuthors,compareYears,compareAuthors,number,norm,validQuartile,scopusRecordUrl,scopusCitedByUrl};
});
