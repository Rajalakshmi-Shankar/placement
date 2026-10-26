labels = sub_label_map()
    if not labels:
        st.info("No submissions yet.")
        return
    sid = st.selectbox("Submission", list(labels), format_func=labels.get)
    history = attempts_for_submission(sid)
    if not history:
        st.info("This submission has not been reviewed yet.")
        return
    aid = st.selectbox("Review attempt", [h["attempt"].id for h in history],
                       format_func=lambda i: f"Review #{i}")
    md, df = build_report(aid)
    st.markdown(md)
    c1, c2 = st.columns(2)
    c1.download_button("Download report (.md)", md, file_name=f"review_{aid}.md", mime="text/markdown")
    c2.download_button("Download criteria (.csv)", df.to_csv(index=False), file_name=f"review_{aid}.csv", mime="text/csv")


def page_candidate():
    st.header("Candidate portal")
    with db() as s:
        cands = s.scalars(select(Candidate).order_by(Candidate.name)).all()
    if not cands:
        st.info("No candidates yet. Ask an admin to import the workbook.")
        return
    st.caption("Prototype: pick yourself from the list (no login).")
    cid = st.selectbox("Who are you?", [c.candidate_id for c in cands],
                       format_func=lambda i: next(f"{c.name} ({c.candidate_id})" for c in cands if c.candidate_id == i))
    with db() as s:
        cand = s.get(Candidate, cid)
        task = s.get(Task, cand.task_id)
        task_d = {"task_id": task.task_id, "title": task.title,
                  "description": task.description, "criteria": task.criteria}
    st.subheader(f"Your task: {task_d['title']}")
    st.write(task_d["description"])
    st.markdown("**Review criteria**\n" + "\n".join(f"- {c}" for c in split_criteria(task_d["criteria"])))

    with st.expander("Task assistant"):
        if st.button("Explain this task"):
            st.session_state[f"assist_{cid}"] = explain_task(task_d)
        if f"assist_{cid}" in st.session_state:
            text, mode = st.session_state[f"assist_{cid}"]
            st.caption(f"Source: {mode}")
            st.markdown(text)

    st.subheader("Submit your GitHub repository")
    with db() as s:
        n_prev = s.scalar(select(func.count(Submission.id)).where(Submission.candidate_id == cid))
    with st.form("submit"):
        url = st.text_input("Public repository URL", placeholder="https://github.com/owner/repo")
        if st.form_submit_button("Submit" if n_prev == 0 else f"Resubmit (attempt {n_prev + 1})"):
            try:
                owner, repo = parse_github_url(url)
                gh_get(f"/repos/{owner}/{repo}")  # confirms the repository exists and is public
                with db() as s:
                    s.add(Submission(candidate_id=cid, task_id=task_d["task_id"],
                                     github_url=f"https://github.com/{owner}/{repo}",
                                     attempt_number=n_prev + 1))
                    s.commit()
                st.success("Submitted. An admin will run the review.")
                st.rerun()
            except RepoError as exc:
                st.error(str(exc))

    st.subheader("Your submissions and feedback")
    with db() as s:
        subs = s.scalars(select(Submission).where(Submission.candidate_id == cid)
                         .order_by(Submission.attempt_number.desc())).all()
        meets = s.scalars(select(Meeting).where(Meeting.candidate_id == cid)).all()
    if not subs:
        st.write("Nothing submitted yet.")
    for sb in subs:
        with st.expander(f"Attempt {sb.attempt_number} - {sb.github_url} - {sb.submitted_at:%Y-%m-%d %H:%M} UTC"):
            hist = attempts_for_submission(sb.id)
            if not hist:
                st.write("Waiting for review.")
            for item in hist:
                show_attempt(item)
    for m in meets:
        st.info(f"SIMULATED meeting: {m.scheduled_for:%Y-%m-%d %H:%M} UTC - {m.note or 'no note'} "
                "(demo only, no real invite was sent)")


# =============================================================================
# 10. MAIN
# =============================================================================
def main():
    st.set_page_config(page_title="AI Review System (prototype)", page_icon="\U0001F50D", layout="wide")
    try:
        get_session_factory()
    except Exception as exc:
        st.error(f"Could not open the database at {DB_PATH}: {exc}")
        st.stop()
    st.sidebar.title("AI Review System")
    st.sidebar.caption("Prototype. No authentication. AI feedback is advisory; humans decide.")
    role = st.sidebar.radio("Role (demo switch)", ["Admin", "Candidate"])
    st.sidebar.write("LLM: " + (f"enabled ({LLM_MODEL})" if llm_enabled() else "**not configured - DEMO mode**"))
    if role == "Admin":
        pages = {"Dashboard": page_dashboard, "Import": page_import, "Review Queue": page_review_queue,
                 "Meetings": page_meetings, "Reports": page_reports}
        choice = st.sidebar.radio("Page", list(pages))
        try:
            pages[choice]()
        except Exception as exc:
            st.error(f"Something went wrong: {exc}")
    else:
        try:
            page_candidate()
        except Exception as exc:
            st.error(f"Something went wrong: {exc}")


main()
