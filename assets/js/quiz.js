(function () {
  function render(rootEl, quiz) {
    rootEl.innerHTML = "";
    const form = document.createElement("form");
    form.className = "quiz";

    quiz.questions.forEach((q, idx) => {
      const block = document.createElement("div");
      block.className = "quiz__question";
      block.dataset.idx = idx;

      const heading = document.createElement("h3");
      heading.textContent = `${idx + 1}. ${q.question}`;
      block.appendChild(heading);

      q.options.forEach((opt, optIdx) => {
        const id = `q${idx}-o${optIdx}`;
        const label = document.createElement("label");
        label.className = "quiz__option";
        label.htmlFor = id;

        const input = document.createElement("input");
        input.type = "radio";
        input.name = `q${idx}`;
        input.id = id;
        input.value = String(optIdx);

        label.appendChild(input);
        label.appendChild(document.createTextNode(opt));
        block.appendChild(label);
      });

      const feedback = document.createElement("div");
      feedback.className = "quiz__feedback";
      feedback.style.display = "none";
      block.appendChild(feedback);

      form.appendChild(block);
    });

    const controls = document.createElement("div");
    controls.className = "quiz__controls";

    const submit = document.createElement("button");
    submit.type = "submit";
    submit.className = "quiz__submit";
    submit.textContent = "Submit answers";

    const reset = document.createElement("button");
    reset.type = "button";
    reset.className = "quiz__reset";
    reset.textContent = "Reset";

    controls.appendChild(submit);
    controls.appendChild(reset);
    form.appendChild(controls);

    const score = document.createElement("div");
    score.className = "quiz__score";
    score.style.display = "none";
    form.appendChild(score);

    rootEl.appendChild(form);

    form.addEventListener("submit", (ev) => {
      ev.preventDefault();
      grade(form, quiz, score);
    });

    reset.addEventListener("click", () => {
      form.reset();
      form.querySelectorAll(".quiz__option").forEach((el) => {
        el.classList.remove("correct", "incorrect");
      });
      form.querySelectorAll(".quiz__feedback").forEach((el) => {
        el.style.display = "none";
        el.textContent = "";
      });
      score.style.display = "none";
      score.className = "quiz__score";
    });
  }

  function grade(form, quiz, scoreEl) {
    let correct = 0;
    quiz.questions.forEach((q, idx) => {
      const block = form.querySelector(`.quiz__question[data-idx="${idx}"]`);
      const selected = block.querySelector(`input[name="q${idx}"]:checked`);
      const labels = block.querySelectorAll(".quiz__option");
      labels.forEach((l) => l.classList.remove("correct", "incorrect"));

      labels.forEach((label, optIdx) => {
        if (optIdx === q.answer) label.classList.add("correct");
      });

      if (selected) {
        const chosen = parseInt(selected.value, 10);
        if (chosen === q.answer) {
          correct += 1;
        } else {
          labels[chosen].classList.add("incorrect");
        }
      }

      const feedback = block.querySelector(".quiz__feedback");
      feedback.textContent = q.explanation || "";
      feedback.style.display = q.explanation ? "block" : "none";
    });

    const total = quiz.questions.length;
    const pct = Math.round((correct / total) * 100);
    const passed = pct >= 70;
    scoreEl.style.display = "block";
    scoreEl.className = "quiz__score " + (passed ? "quiz__score--pass" : "quiz__score--fail");
    scoreEl.textContent = `Score: ${correct} / ${total} (${pct}%) — ${passed ? "well done!" : "review the lesson and try again."}`;
    scoreEl.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  document.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll("[data-quiz]").forEach((rootEl) => {
      const url = rootEl.getAttribute("data-quiz");
      fetch(url)
        .then((r) => {
          if (!r.ok) throw new Error("Could not load quiz: " + url);
          return r.json();
        })
        .then((quiz) => render(rootEl, quiz))
        .catch((err) => {
          rootEl.textContent = "Failed to load quiz: " + err.message;
        });
    });
  });
})();
