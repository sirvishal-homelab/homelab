# 02 - Open questions

| ID | Question | How to answer | Blocks |
|----|----------|---------------|--------|
| Q-1 | Ultra or Pro? Firmware version? | Box label, or the device web page | Everything |
| Q-2 | Stock firmware HTTP API: how to upload an image and select it; size and format limits; does it keep working when left in "photo" mode? | Spike: browse the web UI, capture its requests, upload a test 240x240 JPEG | Approach A |
| Q-3 | Where will it sit (desk, kitchen, rack)? Who looks at it? | Owner | Which idea first |
| Q-4 | Which signals deserve "red"? (e.g. a Proxmox node down, disk > 90 %, Pi-hole down, backup missed) | Owner | Idea 1 |
| Q-5 | Does Beszel expose what we need over its API, and with what credentials (read-only user)? | Read Beszel docs; try with a read-only account | Idea 1 |
| Q-6 | Is Home Assistant on the roadmap? | Owner | Approach B |
| Q-7 | Night mode: dim or off overnight? Does the stock firmware allow brightness control over HTTP? | Spike | Comfort |
| Q-8 | Pi-hole DHCP reservation for the screen (as for the TRMNL, P-1) | Owner, in Pi-hole | Deployment |
