# 02 - Open questions

| ID | Question | How to answer | Blocks |
|----|----------|---------------|--------|
| Q-1 | ~~Ultra or Pro? Firmware version?~~ **Answered 2026-11-14: SmallTV-Ultra, Ultra-V9.0.50.** ESP8266, so ESPHome (approach B) is constrained; approach A is the clear fit | Device web page | - |
| Q-2 | **Mostly answered 2026-11-14, see `03-device-api.md`** (upload + set image verified; theme/brightness strings pending). Stock firmware HTTP API: how to upload an image and select it; size and format limits; does it keep working when left in "photo" mode? | Spike: browse the web UI, capture its requests, upload a test 240x240 JPEG | Approach A |
| Q-3 | **Answered: office, desk or rack; owner only.** Where will it sit (desk, kitchen, rack)? Who looks at it? | Owner | Which idea first |
| Q-4 | Which signals deserve "red"? (e.g. a Proxmox node down, disk > 90 %, Pi-hole down, backup missed) | Owner | Idea 1 |
| Q-5 | **Partly answered: yes for hosts and Docker; owner chose Beszel only: 3 focus systems + container and system counts (see `04`). Owner will create an app account.** Does Beszel expose what we need over its API, and with what credentials (read-only user)? | Read Beszel docs; try with a read-only account | Idea 1 |
| Q-6 | Is Home Assistant on the roadmap? (Not needed for the chosen direction.) | Owner | Approach B |
| Q-7 | Night mode: dim or off overnight? Does the stock firmware allow brightness control over HTTP? | Spike | Comfort |
| Q-8 | Pi-hole DHCP reservation for the screen (as for the TRMNL, P-1) | Owner, in Pi-hole | Deployment |
