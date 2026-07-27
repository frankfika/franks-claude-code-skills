---
name: github-recommend-wechat
description: Turn someone else’s GitHub repository into a source-backed Chinese WeChat recommendation article and publishable HTML/ZIP bundle. Use when the user asks to introduce, recommend, review, or write about a third-party GitHub or open-source project that the user did not create, especially when it must remain separate from the user’s personal Vibe Coding series.
---

# GitHub Project Recommendation → WeChat

Write an independent recommendation, not a project announcement and not an author’s retrospective. Use `wechat-silicon-editor` for drafting checks, typesetting, preview, packaging, and archive delivery.

## Protect the identity boundary

Treat the repository owner and contributors as third parties unless the user explicitly says otherwise.

- Never imply that the user made, launched, maintains, or represents the project.
- Never write “我做了”“我们发布了”“我的项目” or equivalent ownership language.
- First-person observation is welcome: “最近看到”“我比较喜欢”“我不会把它推荐成……”.
- Attribute repository claims to the project when they are not independently verified.
- Do not speak on behalf of the maintainers or promise support, roadmaps, stability, or commercial suitability.

Keep this content outside the user’s self-built-project series:

- Do not use `Vibe Coding｜` or another established personal series label unless the user explicitly requests it.
- Do not explain the distinction with an awkward title such as “这次不是我做的”.
- Use an independent, natural title such as:
  - `推荐一个把小说变成 AI 短剧的开源工作台：Toonflow`
  - `这个开源工具，把零散的 AI 视频步骤串成了一条工作流`

## Verify the repository

Read the current repository before writing:

- README and linked documentation;
- default branch, latest commit, current release, and supported platforms;
- license and any separate frontend/backend repositories;
- official screenshots, demo media, and installation requirements;
- model APIs, paid services, hardware, signing, or deployment prerequisites.

Verify unstable facts from GitHub or another primary source. Do not call a project open source, free, production-ready, self-hostable, or commercially usable unless the current license and repository support that wording. Treat performance, cost, and time-saving figures in the README as project claims unless independently reproduced.

## Write the recommendation

Build one clear argument around why the project is useful and where its limits are. Prefer this reading flow:

1. Open with the practical problem and how the project came to the writer’s attention.
2. Explain the product with a concrete analogy and its main differentiator.
3. Walk through the few workflows that let readers understand how it is used. For each, explain what happens, who benefits, and why it matters.
4. State the real setup, API, cost, platform, or maintenance requirements.
5. Explain who should consider it and who may find it too heavy.
6. End with the repository link.

Write in relaxed, direct Chinese with varied paragraph lengths and a mildly opinionated observer’s voice. Avoid feature dumps, startup-deck language, generic industry analysis, and conclusions that merely repeat the opening.

## Keep the recommendation neutral

- Do not ask readers to Star, follow, join a community, buy, register, or contact the maintainers.
- Do not act as the project’s promoter or community manager.
- A current Star or fork count may appear as a dated factual signal when useful, but never turn it into a call to action.
- End neutrally with `项目地址：<repository URL>` unless the user requests another ending.

## Use evidence-led images

Use real repository, product, or official-source images in reading order:

1. a current GitHub repository screenshot near the opening;
2. the product home or positioning view;
3. one image for each important workflow;
4. optionally an admin, customization, or deployment image when it proves the differentiator.

Do not generate decorative art, fake dashboards, or filler. Name local images sequentially from `01_`. Put a caption immediately after each image in the form `图 01｜…… 图片来源：……`.

## Build and deliver

Read and follow `wechat-silicon-editor` for the actual article build. Use:

- exactly one `h1`;
- `OPEN SOURCE / PROJECT NOTES` as the default eyebrow;
- `--strict-editorial`;
- local numbered images with complete source captions;
- Markdown, HTML preview, image folder, and ZIP delivery;
- the same stable article archive directory for later revisions.

Rebuild every artifact after title or copy changes. Verify the title, licensing language, image paths, copy button, and repository URL. Tell the user to open the HTML, click `复制公众号正文`, and paste it into WeChat.

When a user corrects the ownership tone, series label, title, or call to action, treat that correction as a durable rule for the article and update every generated artifact and archive copy.
