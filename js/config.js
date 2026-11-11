/* =====================================================================
   CSV Combiner settings.
   This is the one place to change limits or switch the login gate on.
   ===================================================================== */
window.COMBINER_CONFIG = {
  REQUIRE_LOGIN: false,     // flip to true later to gate the tool behind BIM-Press login
  MAX_FILES: 50,
  MAX_FILE_MB: 25,
  OUTPUT_PREFIX: "Combined"
};
