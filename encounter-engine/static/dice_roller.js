// Dice Roller — loaded as a static file to avoid template caching issues

var _drCount = 1;
var _drHistory = [];
var _drBuilt = false;

function openDiceRoller() {
  if (!_drBuilt) {
    var dieSides = [4, 6, 8, 10, 12, 20, 100];
    document.getElementById('dr-die-btns').innerHTML = dieSides.map(function(d) {
      return '<button class="btn dr-die-btn" onclick="diceRollerAdd(' + d + ')">d' + d + '</button>';
    }).join('');
    var counts = [1, 2, 3, 4, 5, 6, 8, 10];
    document.getElementById('dr-count-btns').innerHTML = counts.map(function(n) {
      return '<button class="btn btn-ghost btn-xs dr-count-btn' + (n === 1 ? ' dr-count-active' : '') +
        '" onclick="diceRollerSetCount(' + n + ')">' + n + '</button>';
    }).join('');
    _drBuilt = true;
  }
  document.getElementById('dice-roller-modal').style.display = 'flex';
  setTimeout(function() { document.getElementById('dr-formula').focus(); }, 50);
}

function closeDiceRoller(e) {
  if (e && e.target !== document.getElementById('dice-roller-modal')) return;
  document.getElementById('dice-roller-modal').style.display = 'none';
}

function diceRollerSetCount(n) {
  _drCount = n;
  document.querySelectorAll('.dr-count-btn').forEach(function(b) {
    b.classList.toggle('dr-count-active', parseInt(b.textContent) === n);
  });
}

function diceRollerAdd(sides) {
  var inp = document.getElementById('dr-formula');
  var current = inp.value.trim();
  // Parse existing formula into a counts map so we can increment
  var counts = {};  // "8" -> count
  var flat = 0;
  if (current) {
    var toks = current.toLowerCase().replace(/\s+/g,'').match(/[+-]?[^+-]+/g) || [];
    for (var i = 0; i < toks.length; i++) {
      var tok = toks[i];
      var neg = tok.charAt(0) === '-';
      var part = tok.replace(/^[+-]/,'');
      var dm = part.match(/^(\d+)d(\d+)$/);
      if (dm) {
        var key = dm[2];
        counts[key] = (counts[key] || 0) + (neg ? -parseInt(dm[1]) : parseInt(dm[1]));
      } else if (/^\d+$/.test(part)) {
        flat += parseInt(part) * (neg ? -1 : 1);
      }
    }
  }
  // Add the new dice
  var skey = String(sides);
  counts[skey] = (counts[skey] || 0) + _drCount;
  // Rebuild formula string: dice terms sorted by sides, then flat
  var parts = [];
  Object.keys(counts).sort(function(a,b){return parseInt(a)-parseInt(b);}).forEach(function(k) {
    if (counts[k] !== 0) parts.push(counts[k] + 'd' + k);
  });
  if (flat !== 0) parts.push(flat > 0 ? '+' + flat : String(flat));
  inp.value = parts.join('+').replace(/\+-/g,'-');
}

function diceRollerAddMod() {
  var mod = parseInt(document.getElementById('dr-mod').value || '0');
  if (!mod) return;
  // Reuse diceRollerAdd logic: just manipulate the flat part via a fake parse
  var inp = document.getElementById('dr-formula');
  var current = inp.value.trim();
  var counts = {};
  var flat = mod;
  if (current) {
    var toks = current.toLowerCase().replace(/\s+/g,'').match(/[+-]?[^+-]+/g) || [];
    for (var i = 0; i < toks.length; i++) {
      var tok = toks[i];
      var neg = tok.charAt(0) === '-';
      var part = tok.replace(/^[+-]/,'');
      var dm = part.match(/^(\d+)d(\d+)$/);
      if (dm) {
        var key = dm[2];
        counts[key] = (counts[key] || 0) + (neg ? -parseInt(dm[1]) : parseInt(dm[1]));
      } else if (/^\d+$/.test(part)) {
        flat += parseInt(part) * (neg ? -1 : 1);
      }
    }
  }
  var parts = [];
  Object.keys(counts).sort(function(a,b){return parseInt(a)-parseInt(b);}).forEach(function(k) {
    if (counts[k] !== 0) parts.push(counts[k] + 'd' + k);
  });
  if (flat !== 0) parts.push(flat > 0 ? (parts.length ? '+' : '') + flat : String(flat));
  inp.value = parts.join('+').replace(/\+-/g,'-');
  document.getElementById('dr-mod').value = '0';
}

function diceRollerClear() {
  document.getElementById('dr-formula').value = '';
  document.getElementById('dr-result').innerHTML = '';
}

function diceRollerRoll() {
  var raw = document.getElementById('dr-formula').value.trim();
  if (!raw) return;

  // Split on + or - keeping the sign attached to each token
  var tokens = raw.toLowerCase().replace(/\s+/g, '').match(/[+-]?[^+-]+/g) || [];

  // Accumulate dice by die-type into buckets, collect flat bonus
  var buckets = {};   // key = sides string, value = {sides, neg, count}
  var flatBonus = 0;
  var valid = true;

  for (var i = 0; i < tokens.length; i++) {
    var tok = tokens[i];
    var neg = tok.charAt(0) === '-';
    var part = tok.replace(/^[+-]/, '');
    var dm = part.match(/^(\d+)d(\d+)$/);
    if (dm) {
      var n = parseInt(dm[1]);
      var sides = dm[2];
      if (n < 1 || n > 100 || parseInt(sides) < 2) { valid = false; break; }
      var key = (neg ? '-' : '+') + sides;
      if (!buckets[key]) buckets[key] = { sides: parseInt(sides), neg: neg, count: 0 };
      buckets[key].count += n;
    } else if (/^\d+$/.test(part)) {
      flatBonus += parseInt(part) * (neg ? -1 : 1);
    } else {
      valid = false; break;
    }
  }

  if (!valid) {
    document.getElementById('dr-result').innerHTML =
      '<span style="color:#ff6060">Invalid — use e.g. 2d20+3d4+4</span>';
    return;
  }

  // Roll each bucket and build display
  var grand = 0;
  var parts = [];
  var keys = Object.keys(buckets);
  for (var k = 0; k < keys.length; k++) {
    var b = buckets[keys[k]];
    var rolls = [];
    for (var r = 0; r < b.count; r++) {
      rolls.push(Math.ceil(Math.random() * b.sides));
    }
    var sub = rolls.reduce(function(a, v) { return a + v; }, 0) * (b.neg ? -1 : 1);
    grand += sub;
    parts.push('<span class="dr-term">' + (b.neg ? '-' : '') + b.count + 'd' + b.sides +
      ' [' + rolls.join(',') + ']=' + sub + '</span>');
  }
  if (flatBonus !== 0) {
    grand += flatBonus;
    parts.push('<span class="dr-term">' + (flatBonus >= 0 ? '+' : '') + flatBonus + '</span>');
  }

  document.getElementById('dr-result').innerHTML =
    '<div class="dr-result-box">' +
    parts.join('<span class="dr-plus"> </span>') +
    '<div class="dr-grand-total">= ' + grand + '</div></div>';

  _drHistory.unshift(raw + ' = <strong>' + grand + '</strong>');
  if (_drHistory.length > 5) _drHistory.pop();
  document.getElementById('dr-history').innerHTML =
    _drHistory.map(function(h) { return '<div class="dr-hist-row">' + h + '</div>'; }).join('');

  if (typeof addDiceFeedRaw === 'function') addDiceFeedRaw('Dice: ' + raw + ' = ' + grand);
}
