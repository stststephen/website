const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");

const ROOT = path.join(__dirname, "..");
const XLSX_PATH = path.join(ROOT, "data", "credits.xlsx");
const UPLOADS_DIR = path.join(ROOT, "uploads");
const PORT = process.env.ADMIN_PORT || 4000;

const upload = multer({
  storage: multer.diskStorage({
    destination: UPLOADS_DIR,
    filename: function (req, file, cb) {
      var ext = path.extname(file.originalname).toLowerCase().replace(/[^a-z0-9.]/g, "");
      cb(null, Date.now() + "-" + crypto.randomBytes(4).toString("hex") + ext);
    }
  }),
  limits: { fileSize: 8 * 1024 * 1024 },
  fileFilter: function (req, file, cb) {
    cb(null, /^image\/(png|jpeg|webp|gif)$/.test(file.mimetype));
  }
});

const app = express();
app.use(express.json());
app.use(express.static(__dirname));
app.use("/uploads", express.static(UPLOADS_DIR));

// ---------- Spotify album scraping (public metadata only, no API key needed) ----------

function parseAlbumId(input) {
  var s = String(input || "").trim();
  var m = s.match(/open\.spotify\.com\/album\/([a-zA-Z0-9]+)/);
  if (m) return m[1];
  m = s.match(/^spotify:album:([a-zA-Z0-9]+)$/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9]{15,25}$/.test(s)) return s;
  return null;
}

function metaContent(html, property) {
  var re = new RegExp(property.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + '"\\s+content="([^"]*)"');
  var m = html.match(re);
  return m ? m[1].replace(/&#x27;/g, "'").replace(/&amp;/g, "&").replace(/&quot;/g, '"') : "";
}

function allMatches(html, re) {
  var out = [];
  var m;
  var r = new RegExp(re.source, re.flags.indexOf("g") === -1 ? re.flags + "g" : re.flags);
  while ((m = r.exec(html))) out.push(m[1]);
  return out;
}

async function fetchText(url) {
  var res = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0 (compatible; credits-admin-tool/1.0)" } });
  if (!res.ok) throw new Error("HTTP " + res.status + " fetching " + url);
  return res.text();
}

async function fetchAlbum(albumId) {
  var html = await fetchText("https://open.spotify.com/album/" + albumId);

  var ogTitle = metaContent(html, "og:title");
  var ogDescription = metaContent(html, "og:description");
  var ogImage = metaContent(html, "og:image");

  var album = "", artist = "";
  var titleMatch = ogTitle.match(/^(.*) - (?:Album|Single|EP|Compilation) by (.*?)(?: \| Spotify)?$/i);
  if (titleMatch) {
    album = titleMatch[1];
    artist = titleMatch[2];
  } else {
    album = ogTitle.replace(/\s*\|\s*Spotify$/i, "");
  }
  if (!artist && ogDescription) {
    var descParts = ogDescription.split("·").map(function (s) { return s.trim(); });
    artist = descParts[0] || "";
  }
  var yearMatch = ogDescription.match(/\b(19|20)\d{2}\b/);
  var year = yearMatch ? Number(yearMatch[0]) : "";

  var trackIds = allMatches(html, /music:song" content="https:\/\/open\.spotify\.com\/track\/([a-zA-Z0-9]+)"/);
  if (!trackIds.length) {
    throw new Error("No tracks found on that page — make sure the link is to an album, single, or EP (not a track or playlist).");
  }

  var tracks = await Promise.all(trackIds.map(async function (id) {
    try {
      var res = await fetch("https://open.spotify.com/oembed?url=https://open.spotify.com/track/" + id);
      if (!res.ok) return { song: "" };
      var data = await res.json();
      return { song: data.title || "" };
    } catch (e) {
      return { song: "" };
    }
  }));

  return {
    spotifyUrl: "https://open.spotify.com/album/" + albumId,
    album: album,
    artist: artist,
    year: year,
    artworkUrl: ogImage,
    tracks: tracks
  };
}

async function fetchSpotifyArtwork(spotifyUrl) {
  try {
    var html = await fetchText(spotifyUrl);
    return metaContent(html, "og:image");
  } catch (e) {
    return "";
  }
}

// Guards against ever persisting a same-machine admin-server URL as artwork
// (e.g. from a stale client bug resolving an empty image src to the page's own origin).
async function sanitizeArtworkUrl(artworkUrl, spotifyUrl) {
  var url = String(artworkUrl || "").trim();
  if (!url || !/^https?:\/\/(localhost|127\.0\.0\.1)([:/]|$)/i.test(url)) return url;
  if (spotifyUrl) {
    var real = await fetchSpotifyArtwork(spotifyUrl);
    if (real) return real;
  }
  return "";
}

app.post("/api/fetch-album", async function (req, res) {
  try {
    var albumId = parseAlbumId(req.body.url);
    if (!albumId) {
      return res.status(400).json({ error: "That doesn't look like a Spotify album link. Paste a link like https://open.spotify.com/album/..." });
    }
    var data = await fetchAlbum(albumId);
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

// ---------- social link normalization ----------

function normalizeUrl(input) {
  var s = String(input || "").trim();
  if (!s) return "";
  return /^https?:\/\//i.test(s) ? s : "https://" + s;
}

function normalizeHandle(input, prefix) {
  var s = String(input || "").trim();
  if (!s) return "";
  if (/^https?:\/\//i.test(s)) return s;
  return prefix + s.replace(/^@/, "");
}

app.post("/api/upload-artwork", upload.single("file"), function (req, res) {
  if (!req.file) return res.status(400).json({ error: "No image uploaded, or the file type isn't supported (PNG, JPEG, WEBP, GIF only)." });
  res.json({ url: "uploads/" + req.file.filename });
});

// ---------- reading / writing data/credits.xlsx ----------

var REL_HEADER = ["id", "album", "artist", "year", "bandcamp", "spotify", "youtube", "website", "instagram", "tiktok", "artworkUrl", "bandcampEmbedUrl", "description", "createdAt"];
var TRK_HEADER = ["release_id", "song", "roles", "notes"];
var ARTIST_HEADER = ["artist", "website", "instagram", "tiktok"];
var ARTIST_FIELDS = ["website", "instagram", "tiktok"];

function slugify(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "release";
}

function readData() {
  var wb = XLSX.readFile(XLSX_PATH);
  return {
    wb: wb,
    relRows: XLSX.utils.sheet_to_json(wb.Sheets["Releases"], { defval: "" }),
    trkRows: XLSX.utils.sheet_to_json(wb.Sheets["Tracks"], { defval: "" }),
    artistRows: wb.Sheets["Artists"] ? XLSX.utils.sheet_to_json(wb.Sheets["Artists"], { defval: "" }) : []
  };
}

function writeData(wb, relRows, trkRows, artistRows) {
  var relSheet = XLSX.utils.json_to_sheet(relRows, { header: REL_HEADER });
  var trkSheet = XLSX.utils.json_to_sheet(trkRows, { header: TRK_HEADER });
  relSheet["!cols"] = wb.Sheets["Releases"]["!cols"];
  trkSheet["!cols"] = wb.Sheets["Tracks"]["!cols"];
  wb.Sheets["Releases"] = relSheet;
  wb.Sheets["Tracks"] = trkSheet;
  if (artistRows) {
    var artistSheet = XLSX.utils.json_to_sheet(artistRows, { header: ARTIST_HEADER });
    if (wb.Sheets["Artists"] && wb.Sheets["Artists"]["!cols"]) artistSheet["!cols"] = wb.Sheets["Artists"]["!cols"];
    if (!wb.Sheets["Artists"]) wb.SheetNames.push("Artists");
    wb.Sheets["Artists"] = artistSheet;
  }
  XLSX.writeFile(wb, XLSX_PATH);
}

// Keeps the Artists directory and this artist's other releases in sync with
// whatever social/website value is present, in either direction:
//  - a blank field on this release is filled in from the Artists directory
//  - a filled-in field the directory doesn't have yet is learned and pushed
//    out to the artist's other releases that are still blank for it
function syncArtistLinks(artistRows, relRows, artistName, fields) {
  var name = String(artistName || "").trim();
  if (!name) return;
  var key = name.toLowerCase();
  var artistRow = artistRows.find(function (a) { return String(a.artist || "").trim().toLowerCase() === key; });
  if (!artistRow) {
    artistRow = { artist: name, website: "", instagram: "", tiktok: "" };
    artistRows.push(artistRow);
  }

  var learned = [];
  ARTIST_FIELDS.forEach(function (f) {
    if (fields[f] && !artistRow[f]) {
      artistRow[f] = fields[f];
      learned.push(f);
    } else if (!fields[f] && artistRow[f]) {
      fields[f] = artistRow[f];
    }
  });

  if (learned.length) {
    relRows.forEach(function (r) {
      if (String(r.artist || "").trim().toLowerCase() !== key) return;
      learned.forEach(function (f) { if (!r[f]) r[f] = artistRow[f]; });
    });
  }
}

function releasesFromRows(relRows, trkRows) {
  var tracksByRelease = {};
  trkRows.forEach(function (row) {
    var rid = String(row.release_id || "").trim();
    if (!rid) return;
    var roles = String(row.roles || "").split(",").map(function (s) { return s.trim(); }).filter(Boolean);
    (tracksByRelease[rid] = tracksByRelease[rid] || []).push({
      song: String(row.song || ""),
      roles: roles,
      notes: String(row.notes || "")
    });
  });
  return relRows.map(function (row) {
    var id = String(row.id || "").trim();
    return {
      id: id,
      album: String(row.album || ""),
      artist: String(row.artist || ""),
      year: row.year || "",
      bandcamp: String(row.bandcamp || ""),
      spotify: String(row.spotify || ""),
      youtube: String(row.youtube || ""),
      artworkUrl: String(row.artworkUrl || ""),
      website: String(row.website || ""),
      instagram: String(row.instagram || ""),
      tiktok: String(row.tiktok || ""),
      bandcampEmbedUrl: String(row.bandcampEmbedUrl || ""),
      description: String(row.description || ""),
      createdAt: row.createdAt || 0,
      tracks: tracksByRelease[id] || []
    };
  }).filter(function (r) { return r.id; });
}

async function releaseFieldsFromBody(body) {
  var spotify = String(body.spotify || "");
  return {
    album: String(body.album || "").trim(),
    artist: String(body.artist || "").trim(),
    year: body.year || "",
    bandcamp: String(body.bandcamp || ""),
    spotify: spotify,
    youtube: String(body.youtube || ""),
    website: normalizeUrl(body.website),
    instagram: normalizeHandle(body.instagram, "https://instagram.com/"),
    tiktok: normalizeHandle(body.tiktok, "https://www.tiktok.com/@"),
    artworkUrl: await sanitizeArtworkUrl(body.artworkUrl, spotify),
    bandcampEmbedUrl: String(body.bandcampEmbedUrl || ""),
    description: String(body.description || "")
  };
}

app.post("/api/save-release", async function (req, res) {
  try {
    var body = req.body || {};
    var fields = await releaseFieldsFromBody(body);
    var tracks = Array.isArray(body.tracks) ? body.tracks : [];
    var roles = Array.isArray(body.roles) ? body.roles : [];

    if (!fields.album) return res.status(400).json({ error: "Album title is required." });
    if (!tracks.length) return res.status(400).json({ error: "At least one track is required." });

    var data = readData();
    syncArtistLinks(data.artistRows, data.relRows, fields.artist, fields);

    var existingIds = {};
    data.relRows.forEach(function (r) { existingIds[r.id] = true; });
    var baseId = slugify(fields.album);
    var id = baseId, n = 2;
    while (existingIds[id]) { id = baseId + "-" + n; n++; }

    data.relRows.push(Object.assign({ id: id }, fields, { createdAt: Date.now() }));

    var rolesStr = roles.join(", ");
    tracks.forEach(function (t) {
      data.trkRows.push({
        release_id: id,
        song: String(t.song || ""),
        roles: rolesStr,
        notes: ""
      });
    });

    writeData(data.wb, data.relRows, data.trkRows, data.artistRows);
    res.json({ ok: true, id: id });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.get("/api/releases", function (req, res) {
  try {
    var data = readData();
    var releases = releasesFromRows(data.relRows, data.trkRows);
    releases.sort(function (a, b) { return a.album.localeCompare(b.album); });
    res.json(releases);
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.put("/api/releases/:id", async function (req, res) {
  try {
    var id = req.params.id;
    var body = req.body || {};
    var fields = await releaseFieldsFromBody(body);
    var tracks = Array.isArray(body.tracks) ? body.tracks : [];
    var roles = Array.isArray(body.roles) ? body.roles : [];

    if (!fields.album) return res.status(400).json({ error: "Album title is required." });
    if (!tracks.length) return res.status(400).json({ error: "At least one track is required." });

    var data = readData();
    var relRow = data.relRows.find(function (r) { return r.id === id; });
    if (!relRow) return res.status(404).json({ error: "No release with id \"" + id + "\"." });

    syncArtistLinks(data.artistRows, data.relRows, fields.artist, fields);
    Object.assign(relRow, fields);

    var oldNotesBySong = {};
    data.trkRows.forEach(function (t) {
      if (t.release_id === id) oldNotesBySong[String(t.song || "").toLowerCase()] = t.notes || "";
    });

    var remaining = data.trkRows.filter(function (t) { return t.release_id !== id; });
    var rolesStr = roles.join(", ");
    tracks.forEach(function (t) {
      var song = String(t.song || "");
      remaining.push({
        release_id: id,
        song: song,
        roles: rolesStr,
        notes: oldNotesBySong[song.toLowerCase()] || ""
      });
    });

    writeData(data.wb, data.relRows, remaining, data.artistRows);
    res.json({ ok: true, id: id });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.delete("/api/releases/:id", function (req, res) {
  try {
    var id = req.params.id;
    var data = readData();
    if (!data.relRows.some(function (r) { return r.id === id; })) {
      return res.status(404).json({ error: "No release with id \"" + id + "\"." });
    }
    var relRows = data.relRows.filter(function (r) { return r.id !== id; });
    var trkRows = data.trkRows.filter(function (t) { return t.release_id !== id; });
    writeData(data.wb, relRows, trkRows);
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message || String(err) });
  }
});

app.get("/api/roles", function (req, res) {
  res.sendFile(path.join(ROOT, "data", "roles.json"));
});

app.listen(PORT, function () {
  console.log("Credits admin tool running at http://localhost:" + PORT);
});
