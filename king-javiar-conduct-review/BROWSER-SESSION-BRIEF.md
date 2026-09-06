# Brief for the Fleeta Cloud session — paste this in

Do these in order. Log everything into `Event_Log` and `Clip_Register` in
`Javiar_Vehicle_Conduct_Review_Evidence_Pack.xlsx`. Do not skip step 0.

---

## 0. Before anything — confirm the clock

Open any one event and check the timestamp against a known ServiceM8 stamp
(Sat 5 Sep, 08:34 Parking Impact should sit ~14 minutes after a check-out at
08:20:33 in South Brisbane). If Fleeta is showing UTC or any other offset,
**stop and tell me the offset first** — every event will otherwise be matched
to the wrong leg and the whole timeline is wrong.

## 1. Enumerate every event — Fri 4 Sep and Sat 5 Sep (option 3)

One row per event in `Event_Log`: date, time, type, speed, location/suburb,
GPS if shown. No downloads yet. This gives us the shape before we spend time
pulling files.

Skip 21, 22, 28 and 29 August entirely. Javiar had zero jobs allocated on
28 and 29 Aug in ServiceM8 — the grey days are an unused vehicle, not a
missing camera. That line is closed.

## 2. Trip History — Sat 5 Sep 05:00–08:40 first, then Fri 4 Sep (option 4)

**This is the highest-value pull in the whole job.** ServiceM8 says his travel
clock started at 05:19 for an 08:00 job in Chapel Hill that is a 25-minute
drive from the office, and he checked out at 07:46:38 having been on site
4 seconds without doing the job.

That is 2 hours 27 minutes of travel unaccounted for.

Export the route. Screenshot every stop with its address and dwell time.
I want to know where the vehicle actually was between 05:19 and 07:46.

Then do Fri 4 Sep full day for completeness.

## 3. Download the Saturday cluster — 07:03 to 08:34 (option 1)

All 17 clips. Open the first one and check whether it has an audio track.

> **If there IS cabin audio: STOP.**
> Do not play it, do not scrub it, do not download the rest, do not send it
> to anyone. Note "audio present" in the workbook and tell me immediately.
> Queensland's Invasion of Privacy Act 1971 s43 makes it an offence to use a
> listening device to record a private conversation you are not a party to.
> Until a lawyer clears it, that audio is a liability to us, not evidence
> against him. Video only from that point on.

If there is no audio track, carry on and pull all 17.

## 4. Download every Driving Impact clip — both days (option 2)

These are the collision-adjacent triggers and the most defensible if this ends
in disciplinary action. Hard Cornering and Hard Braking clips are supporting
context; get them if retention allows, but Driving Impact first.

## 5. Hash and log every file

For each file, in `Clip_Register`: ref, source, original filename, event
date-time, who collected it, when, where it is stored, SHA-256.

```
Windows:  certutil -hashfile "<file>" SHA256
Mac:      shasum -a 256 "<file>"
```

Never edit an original. Work on copies. This is what makes the evidence stand
up if he challenges it later.

## 6. Tell me these four things when you're done

1. The timezone check result (step 0).
2. Where the vehicle was 05:19–07:46 on Saturday.
3. Audio: yes or no.
4. Total event counts per day, per type, and the Fleeta retention period on
   our current plan — that sets the deadline for everything still undownloaded.

---

**Do not** contact Javiar, mention the review to any staff member, or share
clips with anyone outside this job. Nothing goes to him in writing until the
passenger question and the audio question are both answered.
