'use strict';
// Fictional sample data so the site is not empty on first run.
// Runs only when the database has no users and DEMO is not "0".
const { db } = require('./db');
const { hashPassword } = require('./lib');
const { uniqueSlug, setTags, activateMembership, createInvoice, getConversation, sendMessage } = require('./core');

const DAY = 24 * 3600 * 1000;
const P = (name, description, signup, monthly, match, date) => ({ name, description, signup, monthly, match, date });

const MATCHMAKERS = [
  {
    person: 'Rivka Adler', email: 'matchmaker@demo.matchstriker.com', business: 'Adler Introductions', city: 'New York, NY', years: 14, approval: 0,
    headline: 'Jewish matchmaking for New Yorkers who are done with the apps',
    communities: ['Jewish', 'Modern Orthodox', 'Conservative'], languages: ['English', 'Hebrew', 'Yiddish'], specialties: ['Marriage-minded', 'Ages 28 to 45', 'Second marriages'],
    bio: 'I grew up in Flatbush watching my grandmother make matches at the kitchen table, and I have been doing it professionally since 2012. I work with Jewish singles across the observance spectrum, from Upper West Side Modern Orthodox to "my mother wants me to marry Jewish". I meet every client in person before I introduce them to anyone.',
    approach: 'We start with a long conversation, usually over coffee. I do not use algorithms. I keep a small roster so I actually know everyone on it, and I only suggest a match when I can tell you why. After each date I call both of you.',
    plans: [P('Free profile', 'Join my roster. I will reach out if you are a fit for one of my clients.', 0, 0, 0, 0), P('Pay per match', 'Active search. You pay only when I introduce you to someone.', 150, 0, 180, 0), P('Full search', 'Priority search, unlimited introductions and date feedback calls.', 400, 220, 0, 0)],
  },
  {
    person: 'Min-ji Park', business: 'Park & Seoul Matchmaking', city: 'Los Angeles, CA', years: 9, approval: 1,
    headline: 'Korean and Korean-American introductions, with family in mind',
    communities: ['Korean', 'Korean-American', 'Christian'], languages: ['Korean', 'English'], specialties: ['Family introductions', 'Professionals', 'Bicoastal'],
    bio: 'I match Korean and Korean-American singles in LA and Orange County. Many of my clients were raised between two cultures, and want a partner who understands both without having to explain.',
    approach: 'I speak with you first, and with your parents too if you would like them involved. Every applicant is interviewed before joining, so everyone you meet has been vetted.',
    plans: [P('Membership', 'Six-month commitment recommended. Includes two curated introductions a month.', 500, 250, 0, 0), P('Introductions only', 'No monthly fee. Pay for each introduction and each date I arrange.', 200, 0, 120, 60)],
  },
  {
    person: 'Sana Khan', business: 'Noor Rishta Services', city: 'Chicago, IL', years: 11, approval: 1,
    headline: 'Halal matchmaking for Muslim professionals in the Midwest',
    communities: ['Muslim', 'South Asian', 'Pakistani'], languages: ['English', 'Urdu', 'Hindi', 'Punjabi'], specialties: ['Wali-inclusive process', 'Divorced and widowed', 'Converts welcome'],
    bio: 'I help Muslim singles and their families find a spouse with shared deen and compatible lives. I have worked with families from Chicago to Detroit to Houston.',
    approach: 'Meetings are arranged with respect for your level of practice, with or without family present. No photos are shared without consent from both sides.',
    plans: [P('Pay per meeting', 'No sign-up fee. Pay only when a meeting is arranged.', 0, 0, 0, 90), P('Family plan', 'For parents searching on behalf of a son or daughter.', 250, 75, 0, 0)],
  },
  {
    person: 'Priya Raman', business: 'Shaadi Sutra', city: 'Edison, NJ', years: 7, approval: 0,
    headline: 'Modern desi matchmaking for the tri-state area',
    communities: ['South Asian', 'Indian', 'Hindu', 'Jain'], languages: ['English', 'Hindi', 'Gujarati', 'Tamil'], specialties: ['NRI', 'Horoscope matching on request', 'Ages 25 to 40'],
    bio: 'A former HR recruiter turned matchmaker. I work with Indian-American professionals who want something between the aunty network and the apps.',
    approach: 'Biodata, but better. I write your profile with you, shortlist matches, and handle first contact with families if you prefer.',
    plans: [P('Starter', 'Profile written and listed with my network.', 0, 0, 100, 0), P('Concierge', 'Hands-on search with weekly check-ins.', 300, 180, 0, 0)],
  },
  {
    person: 'Marcus Bell', business: 'Bell & Bond', city: 'Atlanta, GA', years: 6, approval: 0,
    headline: 'Matchmaking for Black professionals in Atlanta',
    communities: ['Black', 'African American', 'Christian'], languages: ['English'], specialties: ['Executives', 'HBCU alumni', 'Date coaching'],
    bio: 'I started Bell & Bond after introducing eleven couples at my church who are all still married. I focus on busy professionals who want a serious relationship.',
    approach: 'One-on-one intake, a small number of well-chosen introductions, and honest feedback after every date.',
    plans: [P('Per date', 'Simple. I plan the date, you show up.', 0, 0, 0, 75), P('Signature', 'Three months of dedicated search plus coaching.', 600, 200, 0, 0)],
  },
  {
    person: 'Daniela Reyes', business: 'Corazón Matchmaking', city: 'Miami, FL', years: 5, approval: 0,
    headline: 'Bilingual matchmaking for Miami\'s Latino community',
    communities: ['Latino', 'Cuban', 'Colombian', 'Catholic'], languages: ['Spanish', 'English', 'Portuguese'], specialties: ['Bilingual couples', 'Ages 30 to 55'],
    bio: 'Born in Bogotá, raised in Hialeah. I match Spanish and English speakers who want a partner who gets the family, the food and the volume.',
    approach: 'Intake in the language you are most yourself in. I host small dinners every month where clients can meet in person.',
    plans: [P('Community', 'Free. Get invited to my monthly dinners.', 0, 0, 0, 0), P('Personal search', 'One-to-one matching.', 150, 95, 0, 40)],
  },
  {
    person: 'Theo Marsh', business: 'Good Company', city: 'San Francisco, CA', years: 8, approval: 0,
    headline: 'LGBTQ+ matchmaking in the Bay Area',
    communities: ['LGBTQ+', 'Gay', 'Lesbian', 'Queer'], languages: ['English', 'Spanish'], specialties: ['Long-term relationships', 'Ages 30+', 'Newly out'],
    bio: 'I have been matching queer folks in the Bay since 2018. My roster is about half men, half women and nonbinary clients.',
    approach: 'A relaxed first meeting, then one thoughtful introduction at a time. I would rather send you on three great dates than thirty forgettable ones.',
    plans: [P('Pay per match', 'No commitment.', 50, 0, 140, 0), P('Member', 'Ongoing search and quarterly member mixers.', 200, 120, 0, 0)],
  },
  {
    person: 'Shirin Tehrani', business: 'Tehrani Introductions', city: 'Los Angeles, CA', years: 20, approval: 1,
    headline: 'Persian matchmaking from Westwood to Great Neck',
    communities: ['Persian', 'Iranian', 'Jewish', 'Muslim'], languages: ['Farsi', 'English'], specialties: ['Family introductions', 'Bicoastal', 'Second marriages'],
    bio: 'Twenty years of introductions in the Iranian community of Los Angeles, with clients of every faith background.',
    approach: 'I know the families. Discretion is everything, and nothing is shared without your permission.',
    plans: [P('Introduction', 'Pay per introduction once accepted.', 300, 0, 250, 0)],
  },
  {
    person: 'Wei Chen', business: 'Red Thread', city: 'Vancouver, BC', years: 10, approval: 0,
    headline: 'Chinese matchmaking across Vancouver and Toronto',
    communities: ['Chinese', 'Cantonese', 'Taiwanese'], languages: ['Mandarin', 'Cantonese', 'English'], specialties: ['Newcomers', 'Professionals', 'Parents welcome'],
    bio: 'Named for the red thread of fate. I work with Chinese-Canadian singles and recent arrivals building a life here.',
    approach: 'Bilingual intake, verified profiles, and introductions in person rather than endless texting.',
    plans: [P('Listed', 'Free listing in my roster.', 0, 0, 0, 0), P('Guided', 'Active matching with date planning.', 100, 60, 80, 0)],
  },
  {
    person: 'Eleni Pappas', business: 'Astoria Matchmaker', city: 'New York, NY', years: 16, approval: 0,
    headline: 'Greek and Orthodox Christian introductions in Queens',
    communities: ['Greek', 'Orthodox Christian', 'Cypriot'], languages: ['Greek', 'English'], specialties: ['Marriage-minded', 'Ages 25 to 50'],
    bio: 'Everyone in Astoria knows someone I introduced. I match Greek and other Orthodox singles around New York.',
    approach: 'You tell me about your family, I tell you about theirs. Then coffee.',
    plans: [P('Per introduction', 'No fee until I introduce you.', 0, 0, 125, 0)],
  },
  {
    person: 'Hal Whitcombe', business: 'Second Spring', city: 'Boston, MA', years: 12, approval: 0,
    headline: 'Matchmaking for singles over 50 in New England',
    communities: ['Over 50', 'Widowed', 'Divorced'], languages: ['English'], specialties: ['Later-in-life relationships', 'Companionship', 'Retirees'],
    bio: 'After losing my wife I found there was nobody helping people my age meet. So I started. Most of my clients are between 55 and 80.',
    approach: 'Unhurried. We talk about the life you have built and who would fit into it. Daytime dates are popular.',
    plans: [P('Seasonal', 'Three months of introductions.', 0, 85, 0, 0), P('Per date', 'Pay as you go.', 0, 0, 0, 55)],
  },
  {
    person: 'Ngozi Okafor', business: 'Ife Connections', city: 'Houston, TX', years: 6, approval: 1,
    headline: 'Nigerian and West African matchmaking in Texas',
    communities: ['Nigerian', 'West African', 'Igbo', 'Yoruba', 'Christian'], languages: ['English', 'Igbo', 'Yoruba'], specialties: ['Diaspora', 'Healthcare professionals', 'Family introductions'],
    bio: 'I connect Nigerians and other West Africans across Houston and Dallas, many of them first-generation professionals.',
    approach: 'I confirm backgrounds and intentions before any introduction. Families are involved when both of you are ready.',
    plans: [P('Standard', 'Vetted introductions.', 100, 0, 110, 0), P('Premium', 'Dedicated search, both cities.', 350, 150, 0, 0)],
  },
];

const CLIENTS = [
  ['Jordan Levy', 'client@demo.matchstriker.com', 33, 'Brooklyn, NY', 'Product designer. Grew up Conservative, keep a kosher-style home, Friday night dinners matter to me.', 'Someone warm and funny who wants kids and a Jewish home.'],
  ['Talia Rosen', 'talia@example.com', 31, 'Manhattan, NY', 'Pediatric nurse, Modern Orthodox, runs half marathons badly.', 'Shomer Shabbat, kind, ready to settle down.'],
  ['Ari Feldman', 'ari@example.com', 36, 'Hoboken, NJ', 'Attorney. Divorced, no kids. Loves cooking for a crowd.', 'A partner in every sense. Open on observance.'],
  ['Noa Shapiro', 'noa@example.com', 29, 'Queens, NY', 'Israeli-American, works in biotech.', 'Hebrew speaker a plus, sense of humor required.'],
  ['Ben Katz', 'ben@example.com', 41, 'Manhattan, NY', 'Cardiologist. Second time around and wiser for it.', 'Emotionally mature, family-oriented.'],
  ['Maya Gold', 'maya@example.com', 34, 'Brooklyn, NY', 'Teacher and ceramicist.', 'Creative, grounded, wants a family.'],
  ['Sam Weiss', 'sam@example.com', 38, 'Stamford, CT', 'Runs a small architecture firm.', 'Someone curious about the world.'],
];

module.exports = function seed() {
  if (process.env.DEMO === '0' || db.get('SELECT 1 AS x FROM users LIMIT 1')) return;
  const now = Date.now();
  const pw = hashPassword('demo1234');

  db.tx(() => {
    const mmIds = [];
    const plans = {};
    MATCHMAKERS.forEach((m, i) => {
      const email = m.email || `mm${i}@example.com`;
      const uid = db.insert('INSERT INTO users (email, password_hash, name, role, created_at) VALUES (?,?,?,?,?)', email, pw, m.person, 'matchmaker', now - 200 * DAY);
      db.run(
        `INSERT INTO matchmakers (user_id, slug, business_name, headline, bio, approach, city, years, accepting, requires_approval, published, hue, created_at)
         VALUES (?,?,?,?,?,?,?,?,?,?,1,?,?)`,
        uid, uniqueSlug(m.business), m.business, m.headline, m.bio, m.approach, m.city, m.years, i === 7 ? 0 : 1, m.approval, i % 8, now - 200 * DAY);
      setTags(uid, 'community', m.communities);
      setTags(uid, 'language', m.languages);
      setTags(uid, 'specialty', m.specialties);
      plans[uid] = m.plans.map((p) =>
        db.insert('INSERT INTO plans (matchmaker_id, name, description, signup_fee, monthly_fee, per_match_fee, per_date_fee, created_at) VALUES (?,?,?,?,?,?,?,?)',
          uid, p.name, p.description, p.signup * 100, p.monthly * 100, p.match * 100, p.date * 100, now - 200 * DAY));
      mmIds.push(uid);
    });

    const adler = mmIds[0];
    const [free, perMatch, full] = plans[adler];
    const clientIds = CLIENTS.map(([name, email, age, city, about, seeking]) => {
      const uid = db.insert('INSERT INTO users (email, password_hash, name, role, created_at) VALUES (?,?,?,?,?)', email, pw, name, 'client', now - 120 * DAY);
      db.run('INSERT INTO client_profiles (user_id, age, city, about, seeking) VALUES (?,?,?,?,?)', uid, age, city, about, seeking);
      return uid;
    });

    // Adler's roster: [client index, plan, days ago, private note]
    const roster = [
      [0, perMatch, 40, 'Met at Cafe Regular. Very easy to talk to. Wants someone who will do Friday nights.'],
      [1, full, 95, 'Shomer Shabbat is firm. Open to NJ.'],
      [2, perMatch, 70, 'Second marriage. Take it slowly.'],
      [3, full, 28, ''],
      [4, full, 150, 'Schedule is brutal, weeknights only.'],
      [5, free, 12, ''],
    ];
    const ms = roster.map(([ci, plan, ago, notes]) => {
      const at = now - ago * DAY;
      const id = db.insert("INSERT INTO memberships (client_id, matchmaker_id, plan_id, status, notes, created_at) VALUES (?,?,?,'pending',?,?)", clientIds[ci], adler, plan, notes, at);
      const row = db.get('SELECT * FROM memberships WHERE id = ?', id);
      activateMembership(row, { paid: true, at });
      // Catch up paid monthly invoices so the ledger has history, leaving the next bill in the future.
      let next = db.get('SELECT next_bill_at FROM memberships WHERE id = ?', id).next_bill_at;
      while (next && next < now) {
        createInvoice({ membership: row, kind: 'monthly', description: 'Membership: Full search', amount: 22000, paid: true, at: next });
        const d = new Date(next); d.setMonth(d.getMonth() + 1); next = d.getTime();
      }
      db.run('UPDATE memberships SET next_bill_at = ? WHERE id = ?', next, id);
      return row;
    });
    // One pending application (shows the approval-free default plus a paused client).
    db.run("UPDATE memberships SET status = 'paused', next_bill_at = NULL WHERE id = ?", ms[4].id);
    db.run("INSERT INTO memberships (client_id, matchmaker_id, plan_id, status, created_at) VALUES (?,?,?,'pending',?)", clientIds[6], adler, perMatch, now - 1 * DAY);

    // Introductions and dates
    const intro = (a, b, ext, note, status, ago) =>
      db.insert('INSERT INTO introductions (matchmaker_id, a_membership_id, b_membership_id, b_external, note, status, created_at) VALUES (?,?,?,?,?,?,?)',
        adler, a.id, b ? b.id : null, ext, note, status, now - ago * DAY);
    const i1 = intro(ms[0], ms[1], '', 'You both light up talking about your grandparents. Start with a walk in Prospect Park.', 'dating', 21);
    createInvoice({ membership: ms[0], kind: 'match', description: 'Match fee: introduction to Talia', amount: 18000, paid: true, at: now - 21 * DAY });
    db.run('INSERT INTO dates (introduction_id, happens_on, venue, created_at) VALUES (?,?,?,?)', i1, new Date(now - 14 * DAY).toISOString().slice(0, 10), 'Prospect Park, then coffee', now - 15 * DAY);
    db.run('INSERT INTO dates (introduction_id, happens_on, venue, created_at) VALUES (?,?,?,?)', i1, new Date(now - 5 * DAY).toISOString().slice(0, 10), 'Dinner at Miriam', now - 6 * DAY);
    intro(ms[2], ms[3], '', 'Ari cooks, Noa eats. I have a good feeling.', 'introduced', 4);
    createInvoice({ membership: ms[2], kind: 'match', description: 'Match fee: introduction to Noa', amount: 18000, paid: false, at: now - 4 * DAY });
    intro(ms[0], null, 'Dana Mizrahi', 'A friend-of-a-friend outside my roster.', 'closed', 33);
    createInvoice({ membership: ms[0], kind: 'match', description: 'Match fee: introduction to Dana', amount: 18000, paid: true, at: now - 33 * DAY });

    // Demo client is also with a second matchmaker, and has one open invoice.
    const corazon = mmIds[5];
    const m2 = db.insert("INSERT INTO memberships (client_id, matchmaker_id, plan_id, status, created_at) VALUES (?,?,?,'pending',?)", clientIds[0], corazon, plans[corazon][0], now - 9 * DAY);
    activateMembership(db.get('SELECT * FROM memberships WHERE id = ?', m2), { paid: true, at: now - 9 * DAY });
    createInvoice({ membership: ms[0], kind: 'custom', description: 'Date coaching session', amount: 6000, at: now - 2 * DAY });

    // Messages
    const c1 = getConversation(clientIds[0], adler);
    sendMessage(c1, clientIds[0], 'Hi Rivka, thank you for the introduction to Talia. Second date went really well!', now - 4 * DAY);
    sendMessage(c1, adler, 'I heard! She said the same. I will leave you two to it, but call me anytime.', now - 4 * DAY + 3600000);
    sendMessage(c1, clientIds[0], 'Will do. Also, could we talk about the coaching session next week?', now - 3 * 3600000);
    const c2 = getConversation(clientIds[6], adler);
    sendMessage(c2, clientIds[6], 'Hello, I just applied for the pay per match plan. A colleague recommended you.', now - 1 * DAY);

    // Network
    db.run("INSERT INTO connections (from_id, to_id, status, created_at) VALUES (?,?,'accepted',?)", adler, mmIds[9], now - 60 * DAY);
    db.run("INSERT INTO connections (from_id, to_id, status, created_at) VALUES (?,?,'pending',?)", mmIds[7], adler, now - 2 * DAY);
    const c3 = getConversation(adler, mmIds[9]);
    sendMessage(c3, mmIds[9], 'Rivka, I have a lovely Greek-Jewish client from Thessaloniki who may suit your roster better than mine. Can I send her your way?', now - 6 * DAY);
  });
  console.log('Seeded demo data (set DEMO=0 to start with an empty site).');
};
