# Fail with an actionable message if the committed TLH/TLI wrappers are stale.
# Invoked with -DREGEN_DIR=... -DSOURCE_DIR=...
foreach(name "ExampleWrap.tlh" "ExampleWrap.tli")
    execute_process(
        COMMAND ${CMAKE_COMMAND} -E compare_files "${REGEN_DIR}/${name}" "${SOURCE_DIR}/${name}"
        RESULT_VARIABLE differs)

    if(NOT differs EQUAL 0)
        message(FATAL_ERROR
            "${name} is out of date with Example.idl.\n"
            "Regenerate the committed wrappers on Windows with:\n"
            "  py TlhFilePatch.py <build-dir>/TestInterfaces/TestInterfaces.dir/<config>/Example.tlh TestInterfaces/ExampleWrap.tlh")
    endif()
endforeach()
